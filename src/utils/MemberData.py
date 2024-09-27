import re
from datetime import datetime
from .zipcodes import lookup_postcode

class MemberData:
    # Constructor that handles initialization from both a row or individual arguments
    def __init__(
        self,
        voornaam=None,
        naam=None,
        emailadres=None,
        telefoon=None,
        extra_telefoon=None,
        straat=None,
        huisnummer=None,
        gemeente=None,
        geboortedatum=None,
        ban=None,
        row=None,
    ):
        if row:
            # If a row is provided, initialize from the row list
            self.voornaam = row[0]
            self.naam = row[1]
            self.emailadres = row[2]
            self.telefoon = self.format_phone_number(row[3])
            self.extra_telefoon = self.format_phone_number(row[4])
            self.straat = row[5]
            self.huisnummer = row[6]
            self.gemeente = self.get_gemeente_name(row[7])
            self.geboortedatum = self.format_date(row[8])
            self.ban = self.clean_ban(row[9])

        else:
            # Otherwise, initialize from individual arguments
            self.voornaam = voornaam
            self.naam = naam
            self.emailadres = emailadres
            self.telefoon = telefoon
            self.extra_telefoon = extra_telefoon
            self.straat = straat
            self.huisnummer = huisnummer
            self.gemeente = self.get_gemeente_name(gemeente)
            self.geboortedatum = self.format_date(geboortedatum)
            self.ban = self.clean_ban(ban) if ban else None

    def get_gemeente_name(self, gemeente):
        """
        Replace postcode with corresponding gemeente name using lookup_postcode.

        :param gemeente: str: The postcode or name of the gemeente.
        :return: str: The corresponding gemeente name or the original if not found.
        """
        # Check if gemeente is a string and not a postcode (assuming postcodes are numeric)
        if isinstance(gemeente, str) and gemeente.isnumeric() and len(gemeente) > 0:
            gemeente_name = lookup_postcode(
                gemeente
            )  # Use the imported lookup_postcode function
        else:
            gemeente_name = gemeente  # Use the provided string directly

        # Capitalize the first letter and make the rest lowercase
        return gemeente_name.capitalize() if gemeente_name else ""  # Ensure not None

    def format_phone_number(self, phone):
        """Formats the phone number to '0000 00 00 00'."""
        if not phone:
            return ""
        match = re.match(r"^\+(\d{1,3})\s*", phone)
        if match:
            # Remove the country code and add a leading zero
            phone = re.sub(r"^\+(\d{1,3})\s*", "0", phone).strip()
        # Remove any non-digit characters
        digits = re.sub(r"\D", "", phone)
        # If the length is less than 10, prepend a leading zero
        if len(digits) < 10:
            digits = "0" + digits

        # Format the number to '0000 00 00 00'
        if len(digits) >= 10:  # Ensure there are enough digits
            return f"{digits[:-6]} {digits[-6:-4]} {digits[-4:-2]} {digits[-2:]}"

        return phone  # Return original if less than 10 digits

    def format_date(self, date_input):
        """Converts a date string or datetime to the format 'dd/mm/yyyy'."""
        if isinstance(date_input, datetime):
            # If it's already a datetime object, format it directly
            return date_input.strftime("%d/%m/%Y")

        if not date_input:
            return ""

        # Try parsing and formatting the date string
        try:
            # Assuming the input date could be in 'yyyy-mm-dd' or other common formats
            parsed_date = datetime.strptime(
                date_input, "%Y-%m-%d"
            )  # Adjust format if necessary
            return parsed_date.strftime("%d/%m/%Y")
        except ValueError:
            return date_input  # Return original if parsing fails

    def clean_ban(self, ban):
        """Cleans the ban name by removing unwanted characters."""
        return re.sub(r"\s*\(.*?\)", "", ban).strip()

    def __repr__(self):
        return (
            f"{self.voornaam} {self.naam}, {self.emailadres}, {self.telefoon}, "
            f"{self.extra_telefoon}, {self.straat}, {self.huisnummer}, "
            f"{self.gemeente}, {self.geboortedatum}, {self.ban}"
        )

    def __eq__(self, other):
        """Define equality based on name and email."""
        if not isinstance(other, MemberData):
            return NotImplemented
        return (self.voornaam.lower(), self.naam.lower()) == (
            other.voornaam.lower(),
            other.naam.lower(),
        )

    def __hash__(self):
        """Define a hash function for the MemberData class."""
        return hash((self.voornaam, self.naam))
