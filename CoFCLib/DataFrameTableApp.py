import tkinter as tk
from tkinter import ttk
import pandas as pd

class DataFrameTableApp:
    def __init__(self, root, df):
        self.root = root
        self.root.title("DataFrame Table App")

        # Store the DataFrame
        self.df = df

        # Create a refresh button
        refresh_button = ttk.Button(self.root, text="Refresh", command=self.refresh_dataframe)
        refresh_button.pack()

        # Create a Treeview widget to display the DataFrame as a table
        self.tree = ttk.Treeview(self.root, columns=list(self.df.columns), show='headings')

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)  # Adjust the column width as needed

        self.tree.pack()

        self.load_dataframe()

    def load_dataframe(self):
        # Clear existing data from the Treeview
        for i in self.tree.get_children():
            self.tree.delete(i)

        # Insert the DataFrame data into the Treeview
        for _, row in self.df.iterrows():
            self.tree.insert('', 'end', values=tuple(row))

    def refresh_dataframe(self):
        # Replace this with your logic to refresh the DataFrame
        # Example: updated_data = {'Name': ['David', 'Eve', 'Frank'], 'Age': [40, 45, 50]}
        # self.df = pd.DataFrame(updated_data)

        # Load the updated DataFrame into the Treeview
        self.load_dataframe()

test_data_frame_table_app = 0
if __name__ == "__main__" and test_data_frame_table_app:
    # Example usage:
    # Create a Tkinter window and pass a DataFrame as input
    root = tk.Tk()
    data = {'Name': ['Alice', 'Bob', 'Charlie'], 'Age': [25, 30, 35]}
    df = pd.DataFrame(data)
    app = DataFrameTableApp(root, df)
    root.mainloop()
