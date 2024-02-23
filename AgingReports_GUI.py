import os
import glob
import pandas as pd
from datetime import datetime, timedelta
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QScrollArea, QPushButton, QCheckBox, QMessageBox, \
    QFileDialog


class AccountSelectionWindow(QWidget):
    def __init__(self, account_data, column_mapping):
        super().__init__()
        self.account_data = account_data
        self.column_mapping = column_mapping
        self.selected_accounts = []

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area_content = QWidget()
        scroll_area_layout = QVBoxLayout(scroll_area_content)

        self.account_data['account_code'] = self.account_data.iloc[:, self.column_mapping['account_code']]
        self.account_data['account_name'] = self.account_data.iloc[:, self.column_mapping['account_name']].astype(str)

        unique_accounts = self.account_data[['account_code', 'account_name']].drop_duplicates().sort_values(
            by='account_code')
        self.account_checkboxes = []

        for code, name in unique_accounts.values:
            checkbox = QCheckBox(f"{code} - {name}")
            scroll_area_layout.addWidget(checkbox)
            self.account_checkboxes.append((code, checkbox))

        scroll_area.setWidget(scroll_area_content)

        select_button = QPushButton("Select")
        select_button.clicked.connect(self.select_accounts)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)

        layout.addWidget(scroll_area)
        layout.addWidget(select_button)
        layout.addWidget(cancel_button)

        self.setLayout(layout)
        self.setWindowTitle("Select Accounts")


    def select_accounts(self):
        self.selected_accounts = [str(code) for code, checkbox in self.account_checkboxes if checkbox.isChecked()]
        QMessageBox.information(self, "Selection Complete", "Account selection complete.")
        self.close()


def pick_folder():
    """Open a dialog for folder selection and return the selected folder path."""
    app = QApplication([])
    folder_path = QFileDialog.getExistingDirectory(None, "Select Folder", "")
    app.quit()
    return folder_path


def load_excel_files(folder_path):
    """Load all excel files in the specified folder."""
    os.chdir(folder_path)
    excel_files = glob.glob('*.xls*')
    data_frames = []
    for file in excel_files:
        data_frames.append(pd.read_excel(file))
    return pd.concat(data_frames)


def get_column_mapping(columns):
    """Ask user to map required columns."""
    print("Available columns:")
    for i, col in enumerate(columns):
        print(f"{i + 1}. {col}")
    mapping = {}
    mapping['vendor_name'] = int(input("Enter column number for vendor name: ")) - 1
    mapping['entry_date'] = int(input("Enter column number for entry date: ")) - 1
    mapping['debit_amount'] = int(input("Enter column number for debit amount: ")) - 1
    mapping['credit_amount'] = int(input("Enter column number for credit amount: ")) - 1
    mapping['account_name'] = int(input("Enter column number for account name: ")) - 1
    mapping['account_code'] = int(input("Enter column number for account code: ")) - 1
    return mapping


def calculate_age(entry_date, report_date):
    """Calculate age based on entry date and report date."""
    age = report_date - entry_date
    years = age.days // 365
    age_months = age.days % 365 // 30
    return years, age_months


def generate_aging_report(df, column_mapping, report_date, selected_accounts):
    """Generate aging report for selected accounts."""
    # Filter DataFrame based on selected accounts
    df['debit_amount'] = df.iloc[:, column_mapping['debit_amount']].fillna(0)
    df['credit_amount'] = df.iloc[:, column_mapping['credit_amount']].fillna(0)
    df['amount'] = df['debit_amount'] - df['credit_amount']
    df['vendor_name'] = df.iloc[:, column_mapping['vendor_name']]
    df['account_name'] = df.iloc[:, column_mapping['account_name']]
    df['account_code'] = df.iloc[:, column_mapping['account_code']].astype(str)

    df = df[
        df['account_code'].isin(selected_accounts)].copy()  # Create a copy to avoid modifying the original DataFrame

    # Calculate age for each entry
    df.loc[:, 'entry_date'] = pd.to_datetime(df.iloc[:, column_mapping['entry_date']])
    df['age'] = (report_date - df['entry_date']).dt.days

    # Sort transactions by entry date within each vendor group
    df.sort_values(by=['vendor_name', 'entry_date'], inplace=True)

    # Define aging buckets
    buckets = [
        ('Within 1 month', 0, 30),
        ('1-3 months', 31, 90),
        ('3-6 months', 91, 180),
        ('6-12 months', 181, 365),
        ('1-2 years', 366, 730),
        ('2-3 years', 731, 1095),
        ('Over 3 years', 1096, float('inf'))
    ]

    # Create DataFrame to store aging report
    aging_report = pd.DataFrame(
        columns=['Serial No.', 'Vendor Name', 'Account Code', 'Account Name'] + [bucket[0] for bucket in buckets] + [
            'Total Amount'])

    # Iterate over vendors
    serial_no = 1
    for vendor_name, vendor_data in df.groupby('vendor_name'):
        # Initialize row for the vendor in the aging report
        row = {
            'Serial No.': serial_no,
            'Vendor Name': vendor_name,
            'Account Code': ', '.join(vendor_data['account_code'].unique()),  # Aggregate all unique account codes
            'Account Name': ', '.join(vendor_data.groupby('account_code')['account_name'].first())
            # Aggregate corresponding account names
        }

        # Initialize balance for each account
        balance = {acc: 0 for acc in selected_accounts}

        # Iterate over transactions for the vendor
        for _, transaction in vendor_data.iterrows():
            account_code = transaction['account_code']
            amount = transaction['amount']

            # Update balance based on transaction amount
            balance[account_code] += amount

            # Calculate amounts in respective aging buckets
            for bucket_name, min_days, max_days in buckets:
                amount_in_bucket = sum(
                    val for key, val in balance.items() if min_days <= transaction['age'] <= max_days)
                row[bucket_name] = amount_in_bucket

        # Calculate total amount
        row['Total Amount'] = sum(balance.values())

        # Add row to aging report
        aging_report = aging_report._append(row, ignore_index=True)
        serial_no += 1

    return aging_report


def main():
    folder_path = pick_folder()
    df = load_excel_files(folder_path)
    report_date = pd.to_datetime(input("Enter the report date (YYYY-MM-DD): "))
    column_mapping = get_column_mapping(df.columns)

    # Open account selection window
    app = QApplication([])
    account_selection_window = AccountSelectionWindow(df, column_mapping)
    account_selection_window.show()
    app.exec_()

    # Get selected accounts
    selected_accounts = account_selection_window.selected_accounts

    if not selected_accounts:
        print("No accounts selected. Exiting.")
        return

    print(selected_accounts)
    aging_report = generate_aging_report(df, column_mapping, report_date, selected_accounts)
    current_datetime = datetime.now().strftime('%Y%m%d%H%M%S')  # Get current date and time
    report_name = f"aging_report_{current_datetime}.xlsx"
    report_path = os.path.join(folder_path, report_name)
    aging_report.to_excel(report_path, index=False)
    print(f"Aging report generated: {report_path}")


if __name__ == "__main__":
    main()
