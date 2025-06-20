import sys
import tkinter as tk
from roster_searcher import RosterSearcher
from gui.app import RosterSearchApp

def launch_gui():
    root = tk.Tk()
    app = RosterSearchApp(root)
    root.mainloop()

def main():
    searcher = RosterSearcher()
    print("Excel Roster Search Program")
    print("=" * 50)
    file_path = input("Enter the path to your Excel file (local path or URL): ").strip()
    if not file_path:
        print("No file path provided. Exiting.")
        return
    search_name = input("Enter the name to search for: ").strip()
    password = input("Enter password for Excel file (press Enter if no password): ").strip() or None
    results = searcher.search_person_schedule(file_path, search_name, password)
    searcher.display_results(results)

if __name__ == "__main__":
    # To launch GUI, uncomment the next line:
    launch_gui()
    # To launch CLI, comment the above and uncomment below:
    # main()
