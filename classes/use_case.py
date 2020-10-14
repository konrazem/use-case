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


    def getTables(self):
        # create table from the given row
        table = {}
        tables = []
        appendix = []
        for use_case in range(2, self.maxrow + 1):
            for col in range(1, self.maxcol + 1): # -1 as it is Appendix
                key = self.sheet.cell(row=1, column=col).value
                val = self.sheet.cell(row=use_case, column=col).value
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

                if key != 'Appendix':
                    table[key] = val

                # the last is Appendix
                if key == 'Appendix':
                    if val == 1:
                        # table copy to create new reference to table and not the same!
                        appendix.append(table.copy())
                    if val != 1:
                        tables.append(table.copy())
        return {'tables': tables, 'appendix': appendix}



    def getDescriptions(self):
        desc = {}
        # in UC - DESCRIPTION there is a sheet with categories Name and Description
        # return description for the given category 
        # 1. Go row by row to get Dictionary
        for row in range(2, self.maxrow): # 1 to ommit names 
            key = self.sheet.cell(row=row, column=1).value
            val = self.sheet.cell(row=row, column=2).value
            desc[key] = val
        
        # we have structure:
        # { 'BANK ACCOUNT MANAGEMENT': '...', 'ACCOUNT MANAGEMENT': ' ... ', ... }
        return desc

