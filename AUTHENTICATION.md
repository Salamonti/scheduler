# Authentication for YRH IM Schedule Generator

The application now includes a login system to restrict access to authorized users only.

## Default Credentials

For initial setup, the following default credentials are available:

- **Username**: `admin`
- **Password**: `admin`

**⚠️ IMPORTANT**: Change these credentials before deploying to production.

## Setting Up Secure Credentials

### Option 1: Streamlit Cloud Secrets (Recommended for Streamlit Cloud)

If deploying on Streamlit Community Cloud, create a `.streamlit/secrets.toml` file with:

```toml
[auth]
admin = "your-secure-password"
john_doe = "another-password"
```

You can add multiple username/password pairs as shown.

### Option 2: Environment Variables (for self-hosted deployments)

Set the following environment variables:

```bash
export SCHEDULER_USERNAME="your-username"
export SCHEDULER_PASSWORD="your-password"
```

### Option 3: Hardcoded Fallback (Development Only)

If neither secrets nor environment variables are set, the default `admin/admin` credentials remain active (with a warning).

## How Authentication Works

1. **Login Screen**: When accessing the app, users are presented with a login form.
2. **Session Persistence**: Once logged in, authentication persists for the duration of the browser session (via Streamlit's session state).
3. **Logout**: A logout button is available in the sidebar to end the session.
4. **Access Control**: All schedule data remains hidden until successful authentication.

## Changing Credentials

To change credentials:

1. **Streamlit Cloud**: Update the `secrets.toml` file in the Streamlit dashboard.
2. **Self‑hosted**: Update environment variables and restart the Streamlit server.
3. **Disable Defaults**: To disable the `admin/admin` fallback, remove or comment out the corresponding line in `scheduler_app.py` (look for `if username == "admin" and password == "admin":`).

## Security Notes

- Passwords are stored in plain text in secrets/environment variables. For higher security, consider integrating with a proper identity provider (e.g., OAuth, Active Directory).
- The application does not currently support password hashing. If you need this, please request an enhancement.
- Ensure your deployment environment (Streamlit Cloud, VPS, etc.) is properly secured and not publicly exposed unless intended.

## Troubleshooting

- **"Invalid username or password"**: Verify that the username/password match exactly (case‑sensitive) in your secrets or environment variables.
- **Authentication errors**: Check the Streamlit logs for any exceptions in the `check_credentials` function.
- **Session reset**: Refreshing the page will require re‑authentication (session state is reset).

## Support

For assistance, contact the developer who deployed this application.