from tkinter import ttk


class PointCloudApp:
    def __init__(self, root, point_cloud_handler):
        self.root = root
        self.point_cloud_handler = point_cloud_handler

        self.root.title("Point Cloud Data Viewer")
        self.create_ui()

    def create_ui(self):
        self.tree = ttk.Treeview(self.root)
        self.tree["columns"] = ("X", "Y", "Z")
        self.tree.heading("#0", text="Index")
        self.tree.heading("X", text="X")
        self.tree.heading("Y", text="Y")
        self.tree.heading("Z", text="Z")
        self.tree.pack()

        self.update_treeview()

        self.btn_generate_data = ttk.Button(self.root, text="Generate New Data", command=self.generate_new_data)
        self.btn_generate_data.pack()

        self.btn_import_data = ttk.Button(self.root, text="Import Data from Excel", command=self.import_data_from_excel)
        self.btn_import_data.pack()

        self.btn_save_to_excel = ttk.Button(self.root, text="Save Data to Excel", command=self.save_to_excel)
        self.btn_save_to_excel.pack()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())

        point_cloud_data = self.point_cloud_handler.point_cloud_data

        for index, (x, y, z) in enumerate(point_cloud_data, start=1):
            self.tree.insert("", "end", text=str(index), values=(x, y, z))

    def generate_new_data(self):
        self.point_cloud_handler.generate_synthetic_point_cloud()
        self.update_treeview()

    def import_data_from_excel(self):
        self.point_cloud_handler.import_from_excel()
        self.update_treeview()

    def save_to_excel(self):
        self.point_cloud_handler.save_to_excel()
        self.update_treeview()
