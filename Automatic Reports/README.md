This is a Python script with several functions that produce different reports based on a set of input data. The purpose of these reports is to provide information about defects and other quality control issues related to a specific order.

The join_reports function reads all Excel files in a specified directory and combines them into a single pandas DataFrame. The resulting DataFrame is then used as input for other functions.

The internal_report function takes the DataFrame produced by join_reports and creates a new DataFrame with columns representing various quality control metrics such as lot number, remittance number, gross meters, net meters, and defects at different stages of production. This function returns the new DataFrame.

The client_report function is similar to internal_report, but it produces a DataFrame with different columns and a different format. The purpose of this report is to provide information to the client about the order, including the lot number, remittance number, gross meters, net meters, and total faults.

The order_information function uses data from the input files to produce a dictionary with various pieces of information about the order, such as the material, the family of materials, the client, the color tone and intensity, the color number, the theoretical width, and the gross and net meters.

The report_layout function takes the order information and report data produced by the other functions and uses them to create a formatted PDF report with a header, table, and footer. This function returns the filename of the PDF report.

Overall, this script reads data from multiple Excel files, combines them into a single DataFrame, calculates various quality control metrics, and produces formatted PDF reports that provide important information to internal and external stakeholders.
