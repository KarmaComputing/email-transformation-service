from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    session,
    redirect,
    url_for,
)  # noqa: E501
from requests_oauthlib import OAuth2Session
import imaplib
import socket
from email import policy
from email.parser import BytesParser
from datetime import datetime
from dataclasses import dataclass
from typing import Optional
from dotenv import dotenv_values
import background


config = dotenv_values(".env")

client_id = config["OAUTH_MICROSOFT_ENTRE_CLIENT_ID"]
client_secret = config["OAUTH_MICROSOFT_ENTRE_CLIENT_SECRET"]
redirect_uri = "https://localhost:5000/oauth/callback"
authorization_base_url = (
    "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"  # noqa: E501
)
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
scope = [
    "https://outlook.office365.com/IMAP.AccessAsUser.All",
    "offline_access",
]  # noqa: E501

app = Flask(__name__)
app.config["SECRET_KEY"] = config["SECRET_KEY"]


@dataclass
class EmailSearchSettings:
    email_host: str
    email_user: str
    email_password: str
    imap_search_subject: Optional[str] = None
    imap_search_unseen: Optional[str] = None
    imap_search_since_date: Optional[str] = None


class EmailSearchError(Exception):
    pass


@background.task
def listen_for_emails(settings: EmailSearchSettings, session_context: dict):
    """
    In the background, keep checking for
    emails which match criteria.

    settings: Email login details and search settings
    session_context: Dict. Copy of flask session (dict(session))
    """
    import time

    print("Listening for emails")
    while True:
        print("Check emails")
        raw_emails = do_search_email(
            session_context,
            settings,
            return_raw_email=True,
            outlook_auth=True,
        )  # noqa: E501
        print(f"Found: {raw_emails}")
        time.sleep(1)


@app.route("/connect/outlook")
def connect_outlook():
    oauth = OAuth2Session(client_id, redirect_uri=redirect_uri, scope=scope)
    authorization_url, state = oauth.authorization_url(authorization_base_url)
    session["oauth_state"] = state
    return redirect(authorization_url)


@app.route("/oauth/callback")
def oauth_callback():
    oauth = OAuth2Session(
        client_id, redirect_uri=redirect_uri, state=session["oauth_state"]
    )
    token = oauth.fetch_token(
        token_url,
        client_secret=client_secret,
        authorization_response=request.url,  # noqa: E501
    )
    session["oauth_token"] = token
    return redirect(url_for("index"))


@app.route("/")
def index():
    return render_template("index.html", form_data=request.form or {})


@app.route("/search-email", methods=["POST"])
def search_email():
    data = request.get_json()
    sample_email = ""

    # Extract and construct the SINCE date if parts are provided
    since_day = data.get("since_day")
    since_month = data.get("since_month")
    since_year = data.get("since_year")

    imap_search_since_date = None
    if since_day and since_month and since_year:
        try:
            full_date = datetime.strptime(
                f"{since_day}-{since_month}-{since_year}", "%d-%m-%Y"
            )  # noqa: E501
            imap_search_since_date = full_date.strftime(
                "%d-%b-%Y"
            )  # e.g., 01-Feb-2025 # noqa: E501
        except ValueError:
            imap_search_since_date = None

    # Build the settings object
    settings = EmailSearchSettings(
        email_host=data.get("email_host"),
        email_user=data.get("email_user"),
        email_password=data.get("email_password"),
        imap_search_subject=data.get("imap_search_subject"),
        imap_search_unseen=(
            int(data["imap_search_unseen"])
            if data.get("imap_search_unseen")
            else None  # noqa
        ),
        imap_search_since_date=imap_search_since_date,
    )
    print("Starting to always listen for emails")
    listen_for_emails(settings, dict(session))
    try:
        raw_emails = do_search_email(
            dict(session),
            settings,
            return_raw_email=True,
            outlook_auth=True,
        )  # noqa: E501
        # Take the first raw email message so we can display it to the user
        if len(raw_emails) == 0:
            sample_email = "No email found with that search criteria"
        else:
            sample_email = raw_emails[0]

    except EmailSearchError as e:
        print(f"Error, {e}")
        return jsonify({"error": str(e), "sample_email": ""}), 400

    return jsonify({"sample_email": sample_email})


def do_search_email(
    session_context,
    settings: EmailSearchSettings,
    json_output=False,
    return_raw_email=False,
    outlook_auth=False,
):
    """
    session_context: dict, copy of flask session (dict(session))
                     so that we don't have to deal with 'operating
                     outside of application context'.
    """
    try:
        # Attempt connection and login
        if outlook_auth:
            imap = imaplib.IMAP4_SSL(settings.email_host)
            bearer_token = session_context.get("oauth_token")["access_token"]
            auth_string = f"user={settings.email_user}\x01auth=Bearer {bearer_token}\x01\x01"  # noqa: E501
            imap.authenticate("XOAUTH2", lambda x: auth_string.encode())
        else:
            imap = imaplib.IMAP4_SSL(settings.email_host)
            imap.login(settings.email_user, settings.email_password)
        imap.select("Inbox")
    except socket.gaierror:
        raise EmailSearchError(
            f"Unable to resolve host: {settings.email_host}"
        )  # noqa: E501
    except imaplib.IMAP4.error as e:
        raise EmailSearchError(f"IMAP login/select failed: {str(e)}")
    except Exception as e:
        print(f"{e}")
        raise EmailSearchError(f"Unexpected error: {str(e)}")

    # Build search criteria
    search_criteria = []

    if settings.imap_search_subject:
        search_criteria.append(f'SUBJECT "{settings.imap_search_subject}"')

    if settings.imap_search_unseen == "1":
        search_criteria.append("UNSEEN")

    if settings.imap_search_since_date:
        search_criteria.append(f"SINCE {settings.imap_search_since_date}")

    search_query = " ".join(search_criteria) if search_criteria else "ALL"
    try:
        resp, emails = imap.search(None, search_query)
    except Exception as e:
        imap.close()
        return {"error": f"Error while searching emails: {str(e)}"}

    output_separator = "#" * 80
    json_response = [] if json_output else None
    raw_emails = [] if return_raw_email else None

    # for num in emails[0].split():

    # Fetch only first email even if multiple found
    resp, data = imap.fetch(emails[0].split()[0], "(RFC822)")
    msg = BytesParser(policy=policy.default).parsebytes(data[0][1])
    simplest = msg.get_body(preferencelist=("plain", "html"))
    html_email = msg.get_body(preferencelist=("html")).get_content()
    email_body = "".join(simplest.get_content().splitlines(keepends=True))

    if json_output:
        json_response.append({"email_body": email_body})
    else:
        if return_raw_email:
            raw_emails.append(html_email)
        print(email_body)
        print(output_separator)

    imap.close()

    if json_output:
        return json_response
    if return_raw_email:
        return raw_emails
