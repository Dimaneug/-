from tkinter import *
from mydatabase import MyDataBase


def main():
    root = Tk()
    MyDataBase(root)
    root.mainloop()


if __name__ == "__main__":
    main()
