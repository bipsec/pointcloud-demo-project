# Pointcloud-demo-project

A point cloud is a collection of points in a three-dimensional (3D) coordinate system. Each point in the cloud represents a specific position in 3D space and may also contain additional attributes, such as color or intensity.

## Project Description

The main objective of this project is to demonstrate how to work with point cloud data and Microsoft Excel in Python. The project includes the following functionalities:

1. Generating Point Cloud Data: A function is implimented to genearte reandom point cloud data to work with that has X,Y,Z coordinates. After generating the point cloud data it is used as a dataset for this project.
2. Data Saving to Excel: After generating the dataset, it is saved into an Excel file.
3. Data Import: Data is imported from the excel file via a function. The function reads any Excel file existing in the directory.
4. Visualization: With the help of a basic graphical user interface (GUI) using Tkinter to visualize the generated point cloud data.


## Up and Run



```sh
git clone https://github.com/bipsec/pointcloud-demo-project.git
cd pointcloud
```


## Create a virtual env
```sh
 python -m venv env
 source env/bin/activate
 pip install -r requirements.txt
```

```sh
python main.py
```

# Conclusion
The project provides a simple demonstration of working with point cloud data and Microsoft Excel in Python. It showcases how to generate synthetic point clouds, save them to Excel, import data from Excel, and visualize the 3D point cloud. 
