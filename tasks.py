# tasks.py
# Klasa bazowa dla wszystkich zadań
class Zadanie:
    def __init__(self, tytul, status="Do zrobienia"):
        self.tytul = tytul
        self.status = status

    def __repr__(self):
        return f"{self.tytul} ({self.status})"

# Klasa dla zadań do zrobienia
class TodoTask(Zadanie):
    def __init__(self, tytul):
        super().__init__(tytul, "Do zrobienia")

# Klasa dla zadań w trakcie
class InProgressTask(Zadanie):
    def __init__(self, tytul):
        super().__init__(tytul, "W trakcie")

# Klasa dla zadań wykonanych
class DoneTask(Zadanie):
    def __init__(self, tytul):
        super().__init__(tytul, "Wykonane")