import argparse
from utils.MemberData import MemberData
from utils.Bans import Bans
from openpyxl import load_workbook


def main(filepath, bans_to_dump):
    # Load the Excel file
    workbook = load_workbook(filename=filepath)
    sheet = workbook.active

    # Assuming the data starts from the second row, with the first row as headers
    members_list = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        member = MemberData(row=row)
        members_list.append(member)

    # Create a Bans object from the list of members
    bans = Bans(members=members_list)

    # Print the specified bans
    if bans_to_dump:
        for ban in bans_to_dump:
            if ban in bans.bans:
                # Print members in the specified ban
                bans.print_table_for_ban(ban)
            else:
                print(f"Warning: Ban '{ban}' is not recognized!")
    else:
        # If no bans specified, print all
        bans.print_table()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process member data.")
    parser.add_argument(
        "--filepath",
        required=True,
        help="Path to the Excel file containing member data.",
    )
    parser.add_argument(
        "--dump", "-d", nargs="*", help="List of bans to print.", default=[]
    )

    args = parser.parse_args()
    main(args.filepath, args.dump)
