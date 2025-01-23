# data-project-for-an-online-retailer-platform
Project for sales data
1. Introduction
Business Description:
This project simulates an order and customer management system for an online retailer platform named Olist Store which will store every order produced by the platform.
The objective is to efficiently manage orders, customers, and product data using Access(SQL databases), an Excel front-end, and VBA middleware.
Application Overview:
Core Features:
Import orders and customer data from the database into Excel.
Filter and display orders based on their status.
Analyze the distribution of order statuses.
Insert new orders into the database(Access).
Generate sales pivot tables for analysis.
Designed with a user-friendly interface and dynamic interactions, this tool provides a simple yet effective way to manage orders.
2. Database
Database Structure:
The database contains three main tables:
1.	Customers: Stores customer details, including customer ID, city, and state.
Customer ID: The unique identifier for each customer (Primary Key, Text).
Customer City: The city where the customer resides (Short Text).
Customer State: The state where the customer resides (Short Text).

2.	Products: Stores product details, including product ID, name, category, and price.
Order ID: The unique identifier for the order associated with the product (Primary Key, Text).
Product ID: The unique identifier for each product (Text).
Product Category Name English: The category to which the product belongs (Short Text).
Price: The price of the product in the order (Currency).

3.	Orders: Stores order details, including order ID, customer ID, payment amount, payment method, order status, and timestamp.
Order ID: The unique identifier for each order (Foreign Key, Text).
Customer ID: The identifier for the customer who placed the order (Foreign Key, Text).
Payment Value: The total amount paid for the order (Currency).
Payment Type: The payment method used for the order (Short Text).
Order Status: The current status of the order, e.g., pending, shipped, delivered (Short Text).
Order Purchase Timestamp: The date and time when the order was placed (Date/Time).
3. Front-end
Design
The Excel file contains the following key worksheets:
Home page: This page serves as a navigation hub, providing access to key functionalities including data import, querying orders by customer ID and shipping status, sales data analysis, and storing data back into the database. From this page, you can access and utilize all the available features seamlessly.
(All the following sheet has button to return Home page.)
Database: Shows data imported from the Access database. 
Orders: Displays filtered orders dynamically populated by VBA.
orders_status_distribution: Displays statistical results for order statuses.
customer_id_return: Displays customer’s information populated by VBA.
sales_pivot table: Provides aggregated sales data by months, categories and cities.
new_records: New data can be recorded in the 'new_records' sheet and can be inserted into the database by VBA.
Business Use
Users select a specific order status on the "Home page," and relevant data is updated automatically in the "Orders" sheet.
The pivot table supports sales trend analysis.
The Excel front-end streamlines business operations by providing key functionalities: filtering orders by status for tracking pending or shipped orders, querying orders by customer ID to support CRM efforts, analyzing order status distribution for performance insights, and using the sales pivot table to guide inventory and marketing decisions. It ensures dynamic updates via VBA integration with the Access database and simplifies data entry with the "New Records" sheet for storing new orders efficiently.
4. VBA Middleware
Subroutine Overview:
1.	ImportDataFromAccess:
Imports orders and customer data from the Access database into the "Database" sheet.
Supports flexibility through dynamic SQL queries.
2.	FetchOrdersBySelectedStatus:
Filters orders based on user input status and displays results in the "Orders" sheet.
Includes color coding for enhanced visualization.
3.	DisplayOrderStatusCounts:
Computes the count of orders by status and displays them in the "orders_status_distribution" sheet.
4.	DisplayCustomerOrders:
Fetches all orders for a specific customer and displays them in the "customer_id_return" sheet.
5.	InsertDataIntoAccess:
Inserts new orders from the "new_records" sheet into the database.
Code Highlights:
The code is clean and concise, featuring error handling and dynamic file path checks.
5. Conclusion
Scalability:
The current application is ideal for small businesses and can be expanded to:
Add more tables (e.g., inventory management).
Enable real-time data synchronization.
Integrate advanced visualization tools.
Multi-user access can be enabled through cloud-based databases.
GitHub Link:
GitHub Repository URL
6. References
Data Sources:
Data sets from kaggles: Brazilian E-Commerce Public Dataset by Olist
