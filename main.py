import os
import openpyxl

RNV_LOCATION = f'C:{os.path.sep}Users{os.path.sep}Malcolm{os.path.sep}Desktop{os.path.sep}Working Folder{os.path.sep}RNVRec{os.path.sep}'


class Reconciliation:
    def __init__(self):
        self.credits = []
        self.debits = []
        self.matches = []
        self.wb = None
        self.sheet = None

    def open(self):
        self.wb = openpyxl.load_workbook(filename=f'{RNV_LOCATION}RNV.xlsx')
        self.sheet = self.wb['JDEData']
        for row in self.sheet.rows:
            self.add_row(row=row)

    def save(self):
        self.wb.active = self.wb['Matches']
        for match in self.matches:
            debit = match['debit']
            credit = match['credit']
            self.wb.active.append((debit['vendor'], debit['po'], debit['amount'], debit['document'], '', '', credit['vendor'], credit['po'], credit['amount'], credit['document']))
        self.wb.save(f'{RNV_LOCATION}RNV.xlsx')
        self.wb.close()

    def add_row(self, row):
        amount = row[1].value
        doc = row[8].value
        po = row[12].value
        vendor = row[24].value
        try:
            if float(amount) > 0:
                self.debits.append({"vendor": vendor, "po": po, "amount": abs(float(amount)), "document": doc})
            else:
                self.credits.append({"vendor": vendor, "po": po, "amount": abs(float(amount)), "document": doc})
        except ValueError:
            pass

    def reconcile(self):
        for debit in self.debits:
            for credit in self.credits:
                if self.is_match(debit=debit, credit=credit):
                    print(f'Match: {debit} to {credit}')
                    self.matches.append({'debit': debit, 'credit': credit})
                    self.credits.remove(credit)
        print(f'Found {len(self.matches)} matches to reconcile')

    @staticmethod
    def is_match(debit, credit):
        return debit['vendor'] == credit['vendor'] and debit['po'] == credit['po'] and debit['amount'] == credit['amount']


if __name__ == '__main__':
    rec = Reconciliation()
    rec.open()
    rec.reconcile()
    rec.save()

