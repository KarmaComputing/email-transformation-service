# Email Transformation Service

Stuck with an email template you website builder won't let you change automated emails onto a supplier or similar?

Want to search / replace (change) / transform an email body into something else and forward it on?

# Setup

See file `.env.example`. Copy it to `.env` and insert settings (if
using outlook authentication, see `portal.azure.com`). If not
using outlook, `OAUTH_MICROSOFT.*` settings are not needed), and can use
standard imap credential login.
```
cp .env.example .env
```
