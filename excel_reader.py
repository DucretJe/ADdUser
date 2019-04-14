# coding: utf8
import xlrd

class file():
    """ This objects defines the database used by the script, it takes on argument which is the absolute path of the excel file.
    The class has several methods, seeking for the appropriates coordonates for the data it looks for the following titles:
    - Name
    - Surname
    - Title
    - Domain Name
    - Service
    then it creates a dictionnary with the data under those coordonates """

    def __init__(self, file):
        workbook = xlrd.open_workbook(file)
        self.sheet= workbook.sheet_by_index(0)

        self.name = self.search_coordonates("Name")
        self.surname = self.search_coordonates("Surname")
        self.title = self.search_coordonates("Title")
        self.domain_name = self.search_coordonates("Domain Name")
        self.service = self.search_coordonates("Service")
        self.error = []

        self.dictionnary = self.dict_create()

    def search_coordonates(self, target):
        """This method search for the provided header in the file and returns the coordonnates of the header himself"""
        for row in range(self.sheet.nrows):
            for col in range(self.sheet.ncols):
                if self.sheet.cell_value(row, col) == target:
                    return row, col

    def search_limit(self,target):
        """This method search the end of the provided data
        If it reach an empty cell it calculates the amount of provided data (number of lines). If not it passes to the next line"""
        x = target[0] + 1
        y = target[1]
        init_x = target[0] + 1

        loop = True
        while loop == True:
            try:
                if self.sheet.cell_value(x, y) == '':
                    loop = False
                else:
                    x += 1
            except:
                break

        amount = x - init_x
        return amount

    def dict_builder(self, category, header):
        """Creation loop """

        cat_dict = []
        cat_limit = self.search_limit(category)
        x = category[0] + 1
        y = category[1]

        i = 1
        while i <= int(cat_limit):
            cat = self.sheet.cell_value(x,y)
            x += 1
            cat_dict.append({header : cat})
            i += 1
        return cat_dict

    def dict_create(self):
        """This method creates a dictionnary with the data provided in the excel file."""

        dictionnary = {}
        ### Building Names
        name = self.dict_builder(self.name, 'name')
        surname= self.dict_builder(self.surname, 'surname')
        title= self.dict_builder(self.title, 'title')
        domain= self.dict_builder(self.domain_name, 'domain')
        service= self.dict_builder(self.service, 'service')


        max = len(name)
        i = 0
        while i < max:
            try:
                dictionnary[i] = {**name[i],**surname[i],**title[i],**domain[i],**service[i]}
            except IndexError:
                self.error.append('{} incomplete, ignored.'.format(name[i]))
            i += 1

        return dictionnary


test = file("Classeur1.xlsx")
