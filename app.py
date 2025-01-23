import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tabulate import tabulate  # Install using: pip install tabulate

SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]


def fetch_events_by_date_range(service, start_date, end_date):
    """Fetches events from the user's primary calendar between two specific dates."""
    # Convert start_date and end_date (dd/mm/yyyy) to datetime objects
    start_datetime = datetime.datetime.strptime(start_date, "%d/%m/%Y")
    end_datetime = datetime.datetime.strptime(end_date, "%d/%m/%Y")

    # Set the time to midnight (start of the day) for start_date
    start_datetime = start_datetime.replace(hour=0, minute=0, second=0, microsecond=0)

    # Set the end time to 23:59:59 (end of the day) for end_date
    end_datetime = end_datetime.replace(hour=23, minute=59, second=59, microsecond=0)

    # Convert both to ISO format with "Z" for UTC time zone
    time_min = start_datetime.isoformat() + "Z"
    time_max = end_datetime.isoformat() + "Z"

    # Fetch events within the time range
    events_result = service.events().list(
        calendarId="primary",
        timeMin=time_min,
        timeMax=time_max,
        maxResults=2500,
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    return events_result.get("items", [])


def fetch_all_events(service):
    """Fetches all events from the user's primary calendar."""
    events_result = service.events().list(
        calendarId="primary",
        maxResults=2500,
        singleEvents=True,
        orderBy="startTime",
    ).execute()
    return events_result.get("items", [])


def display_events(events):
    """Displays events in a tabular format with name, date, time, and location."""
    if not events:
        print("No events found.")
        return

    # Format event data for tabulation
    table = []
    for event in events:
        # Event name
        event_name = event.get("summary", "No Title")

        # Date and time
        start_iso = event["start"].get("dateTime", event["start"].get("date"))
        start_date = datetime.datetime.fromisoformat(start_iso.replace("Z", ""))
        date = start_date.strftime("%d/%m/%Y")
        time = start_date.strftime("%H:%M:%S") if "T" in start_iso else "All Day"

        # Location
        location = event.get("location", "No Location")

        # Append to the table
        table.append([event_name, date, time, location])

    # Display the table
    print(tabulate(table, headers=["Event Name", "Date", "Time", "Location"], tablefmt="grid"))


def main():
    creds = None
    token_path = "token.json"

    # Check if token.json exists
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except ValueError:
            print("Invalid token.json file. Regenerating credentials.")

    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=3000)

        # Save the credentials for the next run
        with open(token_path, "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        # Ask the user whether to fetch all events or events between two dates
        print("Choose an option:")
        print("1. Fetch all events")
        print("2. Fetch events between two dates")
        choice = input("Enter your choice (1 or 2): ")

        if choice == "1":
            # Fetch all events
            print("Fetching all events...")
            events = fetch_all_events(service)
        elif choice == "2":
            # Fetch events for a specific date range
            start_date = input("Enter the start date (dd/mm/yyyy): ")
            end_date = input("Enter the end date (dd/mm/yyyy): ")
            print(f"Fetching events between {start_date} and {end_date}...")
            events = fetch_events_by_date_range(service, start_date, end_date)
        else:
            print("Invalid choice!")
            return

        # Display events
        display_events(events)

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()
