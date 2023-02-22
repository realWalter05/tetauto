
class Person:


    def __init__(self, id, name, street, email):
        self.id = id
        self.name = name
        self.street = street
        self.email = email


    def __str__(self):
        return f"Person: id {self.id}, name {self.name}, street {self.street}, email {self.email}"