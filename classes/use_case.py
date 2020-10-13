class UseCase:
    """Siemanko"""

    def __init__(self, sheet):
        self.sheet = sheet
        self.maxcol = self.getMaxCol()
        self.maxrow = self.getMaxRow()


    def getNames(self):
        res = []
        for col in range(1, self.maxcol):
            res.append(self.sheet.cell(row=1, column=col).value)
        return res

    def getMaxRow(self):
        res = self.sheet.max_row
        # if the max_row is good there will not be None and you return res
        for r in range(1, self.sheet.max_row):
            # first col with Id cannot be None
            val = self.sheet.cell(row=r, column=1).value
            if (val is None):
                res = r - 1  # -1 as this is None cell
                # break as it has no sense to loop fouther
                break
        return res


    def getMaxCol(self):
        res = self.sheet.max_column
        # if the max_column is good there will not be None and you return res
        for c in range(1, self.sheet.max_column):
            # first row with names cannot be None
            val = self.sheet.cell(row=1, column=c).value
            if(val is None):
                res = c - 1 # -1 as this is None cell
                # break as it has no sense to loop fouther
                break
        return res

    def printSheet(self):
        print('Title')
        print(self.sheet.title)
        print('number of columns: ')
        print(self.maxcol)

        print('number of rows: ')
        print(self.maxrow)


    def getTable(self, row_number):
        # create table from the given row
        table = {} 
        for col in range(1, self.maxcol):
            key = self.sheet.cell(row=1, column=col).value
            val = self.sheet.cell(row=row_number, column=col).value
            if val is None:
                val = '-'
                
            if key == 'Priority':
                if val == 1:
                    val = 'MUST'

                if val == 2:
                    val = 'SHOULD'

                if val == 3:
                    val = 'COULD'

                if val == 4:
                    val = 'WOULD'

            table[key] = val
        return table

    def getTables(self):
        tables = []
        for row in range(2, self.maxrow):
            tables.append(self.getTable(row))

        return tables
