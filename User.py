class User:
    def __init__(self, surname, name, second_name):
        self.name = name
        self.surname = surname
        self.second_name = second_name

    def __del__(self):
        class_name = self.__class__.__name__
        print('{} уничтожен'.format(class_name))

    def display(self):
        print(self.surname, ' ', self.name, ' ', self.second_name)

    def __eq__(self, other):
        return self.name == other.name and self.surname == other.surname and self.second_name == other.second_name
