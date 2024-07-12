import os
import sys
from pathlib import Path
import pandas as pd
import re

def extract_numeric_value(total_refund):
    total_refund_str = str(total_refund).lower()
    match = re.search(r"[\d.]+", total_refund_str)
    if match:
        number = float(match.group())
        if 'credit' in total_refund_str:
            return abs(number)
        elif 'debit' in total_refund_str:
            return -abs(number)
        else:
            return number
    return 0.0

def template(entry, target, casino_name, lowercase_transactions_set):
    try:
        target.write(f"\nCasino: {casino_name}\n")
        target.write(f"Player ID: {entry['extplayerid']}\n")
        target.write(f"Screen name: {entry['screen_name']}\n")
        
        if 'player_game_id' in entry and not pd.isnull(entry['player_game_id']):
            target.write(f"Player Game ID: {entry['player_game_id']}\n")

        formatted_stake = "{:.2f}".format(float(entry['stake']))

        if isinstance(entry['transactions'], int):
            transactions_ids = str(entry['transactions']).split('|')
        else:
            transactions_ids = entry['transactions'].split('|')

        if casino_name.lower() in lowercase_transactions_set:
            target.write("Transactions ID(s):\n")
            for transactions_id in transactions_ids:
                target.write(f"{transactions_id}\n")
                
        target.write(f"Amount of stake: {formatted_stake} {entry['currency']} \n")

    except Exception as e:
        print(f"Error processing entry: {e}")

if getattr(sys, 'frozen', False):
    ROOT_DIR = Path(sys.executable).parent
else:
    ROOT_DIR = Path(__file__).resolve().parent

os.chdir(ROOT_DIR)

output_directory = ROOT_DIR / "Refunds"
output_directory.mkdir(exist_ok=True)
print("\nBefore you start, rename the columns as follows (otherwise the program will not function propely):")
print("\n    Casino Name column -> casino_name\n\n    External Player ID -> extplayerid\n\n    Player Game ID (optional, for massive refunds within a SINGLE game) -> player_game_id")
print("\n    Screen name -> screen_name\n\n    Player Currency -> currency\n\n    Transaction IDs (not needed if casinos don't require it)-> transactions\n\n    Player Stake -> stake\n\n    Player Refund -> total_refunds")
print("\n\nDon't forget to save the xslx file after editing too :D")
print("\nPease note that CTRL+V will not paste in a terminal.\nCopy your text as you would normally, but then right click in the terminal to paste.")
os.system('pause')
os.system('cls')
print("                                 ___      ")
print("                               /'___\    ") 
print("      __   ____    _ __    __ /\ \__/    ")
print("    /'__`\/\_ ,`\ /\`'__\/'__`\ \ ,__\   ")
print("   /\  __/\/_/  /_\ \ \//\  __/\ \ \_/   ")
print("   \ \____\ /\____\\\ \_\\\ \____\\\ \_\    ")
print("    \/____/ \/____/ \/_/ \/____/ \/_/\n                                         Evo Madrid 2023 - 2024 v2.4\n\n")
print("=-=-Remember to double check and compare the output with the original file-=-=")
while True:
    file_name = input("\nPlease enter the file name (xlsx): ")
    try:
        if not file_name.endswith('.xlsx'):
            file_name += '.xlsx'
        if Path(ROOT_DIR, file_name).exists():
            break
        else:
            print("File not found or incorrect format")
    except:
        pass

output_encoding = 'utf-8'

name = input("Enter your name: ")
print("\nShortcuts:\nEnter '1' for Service Support Specialist\nEnter '2' for Service Support Team Lead\n")
position = input("Enter your position: ")
reason = input("Enter reason for refund: ")
transactions_needed = input("Do you need transactions ID for any licensee? If yes, please specify all separated by comma [,]: ")
transactions_set = set(transactions_needed.split(','))
lowercase_transactions_set = {transactions.strip().lower() for transactions in transactions_set}

entries_dict = {}

try:
    df = pd.read_excel(file_name)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    sys.exit(1)

for _, row in df.iterrows():
    casino_name = row["casino_name"]

    entry = {
        "extplayerid": row["extplayerid"],
        "screen_name": row["screen_name"],
        "currency": row["currency"],
        "player_game_id": row.get("player_game_id"),
        "stake": row["stake"],
        "total_refund": row["total_refund"], 
        "transactions": row["transactions"]
    }

    if casino_name not in entries_dict:
        entries_dict[casino_name] = []

    entries_dict[casino_name].append(entry)


for casino_name, entries in entries_dict.items():
    output_filename = os.path.join(output_directory, f"{casino_name}.txt")

    with open(output_filename, 'w', encoding='utf-8') as target:
        target.write(f"Dear Casino Team,\n\nWe would like to inform you about the following situation:\n\n")
        target.write(f"{reason}\n")

        credit_entries = [entry for entry in entries if extract_numeric_value(entry['total_refund']) > 0]
        debit_entries = [entry for entry in entries if extract_numeric_value(entry['total_refund']) < 0]

        if credit_entries:
            target.write("\nPlease decide whether you would like to proceed with the following transactions to the players in order to correct the game outcome:\n")
            for entry in credit_entries:
                template(entry, target, casino_name, lowercase_transactions_set)
                formatted_credit_amount = "{:.2f}".format(extract_numeric_value(entry['total_refund']))
                target.write(f"Amount to credit: {formatted_credit_amount} {entry['currency']}\n")

        if debit_entries:
            target.write("\nPlease decide if you would like to debit the following customer(s) in order to correct the game outcome:\n")
            for entry in debit_entries:
                template(entry, target, casino_name, lowercase_transactions_set)
                formatted_debit_amount = "{:.2f}".format(abs(extract_numeric_value(entry['total_refund'])))
                target.write(f"Amount to debit: {formatted_debit_amount} {entry['currency']}\n")

        target.write("\nWe apologize for the inconvenience.\n\nBest regards,\n\n")
        match position:
            case '1':
                position = 'Service Support Specialist'
            case '2':
                position = 'Service Support Team Lead'
            case _:
                position = position
        target.write(f"{name} | {position}")

    print(f"File {output_filename} created for {casino_name} in {output_directory}")
