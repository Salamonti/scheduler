# YRH Internal Medicine Schedule Generator

A Streamlit web application for generating fair, conflict‑free physician schedules for Yarmouth Regional Hospital Department of Internal Medicine.

## Features

- **Fairness‑based scheduling** with weighted scoring and historical balancing
- **Multiple service coverage** (ICU, IM, Hospitalist, Dialysis, CV Clinic, Stress Tests, ECG)
- **Blair‑type rotation rules** for ICU coverage
- **Conflict detection** for vacations, missing coverage, and rule violations
- **Excel export** with color‑coded schedules and fairness reports
- **Google Sheets integration** for loading previous schedule counts
- **Holiday calculation** for Nova Scotia statutory holidays
- **Authentication** to restrict access to authorized personnel

## Quick Start

1. **Install dependencies**:
   ```bash
   pip install streamlit pandas openpyxl
   ```

2. **Run the app**:
   ```bash
   streamlit run scheduler_app.py
   ```

3. **Log in** using the credentials set up in your environment (see [AUTHENTICATION.md](AUTHENTICATION.md)).

4. **Configure doctors, services, vacations, and rotation rules**.

5. **Generate schedules** and export to Excel.

## Authentication

The app now includes a login screen. Before deploying, **set up secure credentials**. See [AUTHENTICATION.md](AUTHENTICATION.md) for detailed instructions.

## Deployment

### Streamlit Cloud (Recommended)

1. Push this repository to GitHub.
2. Connect the repo to [Streamlit Community Cloud](https://streamlit.io/cloud).
3. Add your credentials as [Secrets](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management).

### Self‑hosted (VPS, Docker, etc.)

Set environment variables `SCHEDULER_USERNAME` and `SCHEDULER_PASSWORD` before starting the Streamlit server.

## File Structure

- `scheduler_app.py` – Main application (87 KB, includes scheduling engine and UI)
- `requirements.txt` – Python dependencies
- `AUTHENTICATION.md` – Authentication setup guide
- `scheduler_app_backup.py` – Original version without authentication (backup)

## Scheduling Logic

The scheduler uses a **weighted fairness algorithm** that considers:
- Number of previous assignments per service
- Weekend and night penalties
- Consecutive day limits
- Doctor vacations and service qualifications
- Blair‑rule ICU weeks
- Historical balancing across years

## Support

For issues or feature requests, contact the developer.

---

*Built for Yarmouth Regional Hospital Department of Internal Medicine.*