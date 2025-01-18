# Trivy JSON to Excel Report

This repository contains a Python script that converts Trivy vulnerability scanner JSON reports into well-structured Excel sheets. This is a useful tool for users who want to analyze or share Trivy scan results in a more accessible and visual format, such as Excel, with enhanced styling and formatting options.

---

## Features:
- Converts **Trivy JSON output** into an **Excel file**.
- Creates multiple sheets with detailed data:
  - **Summary**: Contains metadata about the artifact/image.
  - **History**: Lists the image creation history.
  - **Vulnerabilities**: Displays detailed vulnerability information.
  - **Unique Vulnerabilities**: Lists unique vulnerabilities based on package and severity.
  - **Unique-Vulnerability-Statistics**: Provides statistical analysis of unique vulnerabilities.
  - **Overall-Statistics**: Gives statistics on vulnerabilities present in the entire report.

---

## Setup Instructions:

### Prerequisites:
Ensure that you have the following installed:
1. **Python 3.x**
2. **Trivy** (for generating JSON reports)

If you don't have **Trivy** installed, you can install it by following [Trivy's installation instructions](https://github.com/aquasecurity/trivy#installation).

3. **Required Python Libraries**:
   - `pandas` (for data manipulation)
   - `openpyxl` (for creating and manipulating Excel files)
   
   You can install the dependencies using `pip`:
   ```bash
   pip install pandas openpyxl
   ```

### 1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/trivy-json-to-excel-report.git
   cd trivy-json-to-excel-report
   ```

### 2. Generate Trivy JSON report:
   Run a Trivy scan and output the results as JSON.
   ```bash
   trivy image --format json -o trivy_report.json <image_name>
   ```

### 3. Run the conversion script:
   Execute the Python script to convert the Trivy JSON output into an Excel report.
   ```bash
   python convert_trivy_json_to_excel.py
   ```

   The script will ask for the **JSON file name** (e.g., `trivy_report.json`). After entering the file name, it will generate an Excel report named `trivy_report.xlsx`.

---

## Excel Output Details:

When you run the script, it will generate an Excel file with the following sheets:

### 1. **Summary Sheet**:
   This sheet provides metadata about the scanned artifact/image. It includes:
   - **Artifact Name**: Name of the artifact.
   - **Artifact Type**: Type of artifact (e.g., image, container).
   - **OS Family**: The OS family of the image (e.g., Linux).
   - **OS Version**: The OS version.
   - **Image ID**: Unique ID of the image.
   - **Created At**: Image creation timestamp.
   - **Repo Tags**: Repository tags associated with the image.
   - **Repo Digests**: Digests of the image in the repository.
   - **Image Created**: Timestamp of when the image was created.

### 2. **History Sheet**:
   This sheet lists the **image creation history**, such as:
   - **Created At**: Timestamp of when the layer was created.
   - **Is Empty Layer?**: Boolean indicating if the layer is empty.
   - **Created By**: Information about the user who created the layer.

### 3. **Vulnerabilities Sheet**:
   This sheet contains detailed information on vulnerabilities found in the image:
   - **Vulnerability ID**: The ID for the vulnerability.
   - **Package**: The package that has the vulnerability.
   - **Installed Version**: The installed version of the package.
   - **Fixed Version**: The version in which the vulnerability is fixed.
   - **Status**: The status of the vulnerability (e.g., fixed, active).
   - **Severity Level**: Severity of the vulnerability (e.g., CRITICAL, HIGH).
   - **Severity Source**: The source of the severity classification.
   - **Primary URL**: A URL for more information about the vulnerability.
   - **Title**: The title of the vulnerability.
   - **Description**: Description of the vulnerability.
   - **References**: Links to references for the vulnerability.
   - **Published Date**: Date when the vulnerability was published.

### 4. **Unique Vulnerabilities Sheet**:
   This sheet filters and lists unique vulnerabilities based on the package and severity level:
   - **Package**: The affected package.
   - **Installed Version**: Installed version of the package.
   - **Fixed Version**: Fixed version for the vulnerability.
   - **Status**: Vulnerability status.
   - **Severity Level**: The severity of the vulnerability.

### 5. **Unique-Vulnerability-Statistics Sheet**:
   Provides overall statistics for unique vulnerabilities:
   - **Total Vulnerabilities**: Total number of unique vulnerabilities found.
   - **Severity Counts**: Counts of vulnerabilities per severity level (e.g., CRITICAL, HIGH).

### 6. **Overall-Statistics Sheet**:
   Displays statistics for the overall vulnerabilities found across the image:
   - **Total Vulnerabilities**: Total number of vulnerabilities found in the image.
   - **Severity Counts**: Counts for each severity level across all vulnerabilities.

---

## Example Output:

After running the script, you will get an Excel file (e.g., `trivy_report.xlsx`) with the above sheets. The sheet names will reflect the following structure:
- `Summary`
- `History`
- `Vulnerabilities`
- `Unique Vulnerabilities`
- `Unique-Vulnerability-Statistics`
- `Overall-Statistics`

---

## Scope for Improvement:
1. **Enhanced Filtering and Analysis**: Add additional filtering and grouping options (e.g., by package name, severity, etc.).
2. **Support for Other Output Formats**: Extend the script to support converting other formats like CSV or PDF.
3. **Customizable Excel Styling**: Allow users to define custom Excel styling (e.g., colors, fonts, cell formatting).
4. **Performance Improvements**: Optimize processing for very large JSON files, possibly by streaming data instead of loading everything into memory.
5. **Integration with CI/CD**: Integrate the script into CI/CD pipelines to automate vulnerability report generation and analysis.

---

## Contributing:
We welcome contributions to improve and extend this project. If you have any enhancements or fixes, feel free to:
1. Fork this repository
2. Create a new branch for your changes
3. Submit a pull request

---


## Notes:
- This script supports only Trivy JSON outputs in the current format (as of the last Trivy release).
- Make sure your Trivy report is up-to-date to ensure compatibility with the script.
  
--- 
