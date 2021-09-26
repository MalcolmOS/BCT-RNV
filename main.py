import openpyxl
import datetime
from pathlib import Path

RNV_LOCATION = Path(f'C:\\Users\\Malcolm\\Desktop\\Working Folder\\AP\\RNVRec')


class Reconciliation:
    def __init__(self):
        self.credits = []
        self.debits = []
        self.matches = []
        self.wb = None
        self.sheet = None

    def open(self) -> None:
        self.wb = openpyxl.load_workbook(filename=RNV_LOCATION.joinpath('RNV.xlsx'))
        self.sheet = self.wb['Sheet1']
        for row in self.sheet.rows:
            self.add_row(row=row)

    def save(self) -> None:
        self.wb.create_sheet("Matches")
        self.wb.active = self.wb['Matches']
        self.wb.active.append(('Vendor', 'PO', 'Debit Amount', 'Document', '', '', 'Vendor', 'PO', 'Credit Amount', 'Document'))
        for match in self.matches:
            debit = match['debit']
            credit = match['credit']
            self.wb.active.append((debit['vendor'], debit['po'], debit['amount'], debit['document'], '', '', credit['vendor'], credit['po'], credit['amount'], credit['document']))
        self.wb.save(RNV_LOCATION.joinpath('RNV.xlsx'))
        self.wb.close()

    def add_row(self, row) -> None:
        amount = row[1].value
        doc = row[8].value
        po = row[12].value
        vendor = row[24].value
        try:
            if float(amount) > 0:
                self.debits.append({"vendor": vendor, "po": po, "amount": abs(float(amount)), "document": doc})
            else:
                self.credits.append({"vendor": vendor, "po": po, "amount": abs(float(amount)), "document": doc})
        except Exception:
            pass

    def reconcile(self) -> None:
        for debit in self.debits:
            if self.has_match(debit=debit):
                continue

    def has_match(self, debit) -> bool:
        for credit in self.credits:
            if self.is_match(debit=debit, credit=credit):
                print(f'Match: {debit} to {credit}')
                self.matches.append({'debit': debit, 'credit': credit})
                self.credits.remove(credit)
                return True
        return False

    @staticmethod
    def is_match(debit, credit) -> bool:
        return debit['vendor'] == credit['vendor'] and debit['po'] == credit['po'] and debit['amount'] == credit['amount']


if __name__ == '__main__':
    start = datetime.datetime.now().replace(microsecond=0)
    rec = Reconciliation()
    rec.open()
    rec.reconcile()
    rec.save()
    end = datetime.datetime.now().replace(microsecond=0)
    print(f'Found {len(rec.matches)} matches to reconcile in {end-start} seconds')
