from .MemberData import MemberData
from tabulate import tabulate  # Import the tabulate library


class Bans:
    def __init__(self, members=None):
        # Dictionary to store bans with lists of MemberData objects
        self.bans = {
            "Piepedollen": [],
            "Speelvogels": [],
            "Krabbekoningen": [],
            "Knapen": [],
            "Jonghernieuwers": [],
            "Hernieuwers": [],
            "Leiding": [],
            "Ondersteunend lid": []
        }

        # If a list of members is provided, add them to the respective bans
        if members:
            self.add_members(members)

    def add_member(self, member: MemberData):
        """Adds a MemberData object to the appropriate ban category and maintains internal sorting."""
        if member.ban in self.bans:
            if member not in self.bans[member.ban]:
                self.bans[member.ban].append(member)
                # Sort members by Voornaam and then by Naam after adding
                self.bans[member.ban].sort(key=lambda m: (m.voornaam, m.naam))
        else:
            print(f"Warning: {member.ban} is not a recognized ban category!")

    def add_members(self, members: list):
        """Adds a list of MemberData objects to the appropriate ban categories."""
        for member in members:
            self.add_member(member)

    def print_table(self):
        """Prints all members in a formatted table for each ban category."""
        for ban, members in self.bans.items():
            print(f"\nBan: {ban} ({len(members)} members)\n")
            if members:  # Check if there are members in the current ban
                # Create a table from member data
                table = [
                    [
                        member.voornaam,
                        member.naam,
                        member.emailadres,
                        member.telefoon,
                        member.extra_telefoon,
                        member.straat,
                        member.huisnummer,
                        member.gemeente,
                        member.geboortedatum,
                        member.ban,
                    ]
                    for member in members
                ]

                # Print the table with headers
                print(
                    tabulate(
                        table,
                        headers=[
                            "Voornaam",
                            "Naam",
                            "Emailadres",
                            "Telefoon",
                            "Extra Telefoon",
                            "Straat",
                            "Huisnummer",
                            "Gemeente",
                            "Geboortedatum",
                            "Ban",
                        ],
                        tablefmt="pretty",
                    )
                )
            else:
                print("  No members in this ban.")

    def print_table_for_ban(self, ban):
        """Prints members for a specific ban category in a formatted table."""
        members = self.bans.get(ban, [])

        if not members:
            print(f"No members in the {ban} category.")
            return

        # Prepare data for tabulate
        table_data = []
        for member in members:
            table_data.append([
                    member.voornaam,
                    member.naam,
                    member.emailadres,
                    member.telefoon,
                    member.straat,
                    member.huisnummer,
                    member.gemeente,
                    member.geboortedatum
                ])

        # Print header and table using tabulate
        headers = ['Voornaam', 'Naam', 'Emailadres', 'Telefoon', 'Straat', 'Huisnummer', 'Gemeente', 'Geboortedatum']
        print(f"\n=== Members in Ban: {ban} ===")
        print(tabulate(table_data, headers=headers, tablefmt='pretty'))

    def __repr__(self):
        result = ""
        for ban, members in self.bans.items():
            result += f"\nBan: {ban} ({len(members)} members)\n"
            for member in members:
                result += f"  - {member}\n"
        return result
