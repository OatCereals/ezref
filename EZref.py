import os
import sys
from pathlib import Path
import pandas as pd

def template(entry, target, casino_name):
    target.write(f"\n\nCasino: {casino_name}\n")
    target.write(f"Player ID: {entry['extplayerid']}\n")
    target.write(f"Screen name: {entry['screen_name']}\n")
    target.write(f"Player Game ID: {entry['player_game_id']}\n")
    formatted_stake = "{:.2f}".format(float(entry['stake']))
    
    if isinstance(entry['transactions'], int):
        transactions_ids = str(entry['transactions']).split('|')
    else:
        transactions_ids = entry['transactions'].split('|')

    if casino_name.lower() in lowercase_transactions_set:
        target.write("transactions ID(s):\n")
        for transactions_id in transactions_ids:
            target.write(f"{transactions_id}\n")
            
    target.write(f"Amount of stake: {formatted_stake} {entry['currency']} \n")

if getattr(sys, 'frozen', False):
    ROOT_DIR = Path(sys.executable).parent
else:
    ROOT_DIR = Path(__file__).resolve().parent

os.chdir(ROOT_DIR)

output_directory = ROOT_DIR / "Refunds"
output_directory.mkdir(exist_ok=True)

print("                                 ___      ")
print("                               /'___\    ") 
print("      __   ____    _ __    __ /\ \__/    ")
print("    /'__`\/\_ ,`\ /\`'__\/'__`\ \ ,__\   ")
print("   /\  __/\/_/  /_\ \ \//\  __/\ \ \_/   ")
print("   \ \____\ /\____\\\ \_\\\ \____\\\ \_\    ")
print("    \/____/ \/____/ \/_/ \/____/ \/_/                      Evo Madrid 2023 v2.2\n\n\n")

while True:
    file_name = input("Hello, please enter the file name (xlsx): ")
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
position = input("Enter your position: ")
reason = input("Enter reason for refund: ")
transactions_needed = input("Do you need transactions ID for any licensee? If yes, please specify all, case-sensitive and separated by comma [,]: ")
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
        "player_game_id": row["player_game_id"],
        "stake": row["stake"],
        "total_refund": float(row["total_refund"]), 
        "transactions": row["transactions"]
    }

    if casino_name not in entries_dict:
        entries_dict[casino_name] = []

    entries_dict[casino_name].append(entry)


for casino_name, entries in entries_dict.items():
    output_filename = os.path.join(output_directory, f"{casino_name}.txt")

    with open(output_filename, 'w', encoding='utf-8') as target:
        target.write(f"Dear Casino Team,\n\nWe would like to inform you about the following situation:\n\n")
        target.write(f"{reason}")

        credit_entries = [entry for entry in entries if entry['total_refund'] > 0]
        debit_entries = [entry for entry in entries if entry['total_refund'] < 0]

        if credit_entries:
            target.write("\n\nPlease decide whether you would like to proceed with the following transactions to the players in order to correct the game outcome:")
            for entry in credit_entries:
                template(entry, target, casino_name)
                formatted_credit_amount = "{:.2f}".format(entry['total_refund'])
                target.write(f"Amount to credit: {formatted_credit_amount} {entry['currency']}")

        if debit_entries:
            target.write("\n\nPlease decide if you would like to debit the following customer(s) in order to correct the game outcome:")
            for entry in debit_entries:
                template(entry, target, casino_name)
                formatted_debit_amount = "{:.2f}".format(abs(entry['total_refund']))
                target.write(f"Amount to debit: {formatted_debit_amount} {entry['currency']}")

        target.write("\n\nWe apologize for the inconvenience.\n\nBest regards,\n\n")
        target.write(f"{name} | {position}")

    print(f"File {output_filename} created for {casino_name} in {output_directory}")