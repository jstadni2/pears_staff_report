# PEARS Staff Report

The [PEARS](https://www.k-state.edu/oeie/pears/) Staff Report summarizes the PEARS activity of [SNAP-Ed](https://www.fns.usda.gov/snap/snap-ed) staff on a monthly basis. Separate reports are generated for each Illinois SNAP-Ed implementing agency, [Illinois Extension](https://inep.extension.illinois.edu/) and [Chicago Partnership for Health Promotion \(CPHP\)](https://cphp.uic.edu/).

## Installation

The recommended way to install the PEARS Staff Report is through git, which can be downloaded [here](https://git-scm.com/downloads). Once downloaded, run the following command:

```bash
git clone https://github.com/jstadni2/pears_staff_report
```

Alternatively, this repository can be downloaded as a zip file via this link:
[https://github.com/jstadni2/pears_staff_report/zipball/master/](https://github.com/jstadni2/pears_staff_report/zipball/master/)

This repository is designed to run out of the box on a Windows PC using Docker and the [/sample_inputs](https://github.com/jstadni2/pears_staff_report/tree/main/sample_inputs) and [/sample_outputs](https://github.com/jstadni2/pears_staff_report/tree/main/sample_outputs) directories.
To run the script in its current configuration, follow [this link](https://docs.docker.com/desktop/windows/install/) to install Docker Desktop for Windows. 

With Docker Desktop installed, this script can be run simply by double clicking the `run_script.bat` file in your local directory.

The `run_script.bat` file can also be run in Command Prompt by entering the following command with the appropriate path:

```bash
C:\path\to\pears_staff_report\run_script.bat
```

### Setup instructions for SNAP-Ed implementing agencies

The following steps are required to generate the PEARS Staff Report using your organization's PEARS data:
1. Contact [PEARS support](mailto:support@pears.io) to set up an [AWS S3](https://aws.amazon.com/s3/) bucket to store automated PEARS exports.
2. Download the automated PEARS exports. Illinois Extension's method for downloading exports from the S3 is detailed in the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting/blob/6f370389776fb8f88495fbe4e7918c203fd84997/pears_nightly_export_reformatting.py#L9-L45).
3. Set the appropriate input and output paths in `pears_staff_report.py` and `run_script.bat`.
	- The [Input Files](#input-files) and [Output Files](#output-files) sections provide an overview of required and output data files.
	- Copying input files to the build context would enable continued use of Docker and `run_script.bat` with minimal modifications.
	- `pears_staff_report.py` may require additional alterations depending on the staff list format. 
4. Set the usename and password variables in [pears_staff_report.py](https://github.com/jstadni2/pears_staff_report/blob/270de975d41a2fea8a9dd83013ed7b56a9460a74/pears_staff_report.py#L279-L280) using valid Office 365 credentials.	

### Additional setup considerations

- The formatting of PEARS export workbooks changes periodically. The example PEARS exports included in the [/sample_inputs](https://github.com/jstadni2/pears_staff_report/tree/main/sample_inputs) directory are based on workbooks downloaded on 08/12/22.
Modifications to `pears_staff_report.py` may be necessary to run with subsequent PEARS exports.
- Illinois Extension utilized [Task Scheduler](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page) to run this script from a Windows PC on a monthly basis.
- Plans to deploy the PEARS Staff Report on AWS were never implemented and are currently beyond the scope of this repository.
- Other SNAP-Ed implementing agencies intending to utilize the PEARS Staff Report should consider the following adjustments as they pertain to their organization:
	- If your organization actively maintains its SNAP-Ed staff list internally in PEARS, the `User_Export.xlsx` workbook could be used in lieu of external staff lists.
	- The `compile_report()` and `save_staff_report()` functions both require an `agency` input argument to format separate reports for the two Illinois SNAP-Ed implementing agencies. Adjust as needed for your organization's specification.
	- The `send_mail()` function in [pears_staff_report.py](https://github.com/jstadni2/pears_staff_report/blob/270de975d41a2fea8a9dd83013ed7b56a9460a74/pears_staff_report.py#L313) is defined using Office 365 as the host. Change the host to the appropriate email service provider if necessary.

## Input Files

The following input files are required to run the PEARS Staff Report script:
- [FY22_INEP_Staff_List.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/FY22_INEP_Staff_List.xlsx): A workbook that compiles various lists of Illinois Nutrition Education Programs staff.
- Reformatted PEARS module exports output from the [PEARS Nightly Export Reformatting script](https://github.com/jstadni2/pears_nightly_export_reformatting):
	- [Coalition_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/Coalition_Export.xlsx)
	- [Indirect_Activity_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/Indirect_Activity_Export.xlsx)
	- [Partnership_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/Partnership_Export.xlsx)
	- [Program_Activities_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/Program_Activities_Export.xlsx)
	- [PSE_Site_Activity_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/PSE_Site_Activity_Export.xlsx)
	- [Success_Story_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/Success_Story_Export.xlsx)
	- [User_Export.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_inputs/User_Export.xlsx)

Example input files are provided in the [/sample_inputs](https://github.com/jstadni2/pears_staff_report/tree/main/sample_inputs) directory. 
PEARS module exports included as example files are generated using the [Faker](https://faker.readthedocs.io/en/master/) Python package and do not represent actual program evaluation data. 

## Output Files

The following output files are produced by the PEARS Staff Report script:
- [CPHP Staff PEARS Entries YYYY-MM.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_outputs/CPHP%20Staff%20PEARS%20Entries%202022-05.xlsx): A workbook that summarizes the PEARS activity of CPHP staff for the given month and year to date.
- [Extension Staff PEARS Entries YYYY-MM.xlsx](https://github.com/jstadni2/pears_staff_report/blob/main/sample_outputs/Extension%20Staff%20PEARS%20Entries%202022-05.xlsx): A workbook that summarizes the PEARS activity of Illinois Extension staff for the given month and year to date.

Example output files are provided in the [/sample_outputs](https://github.com/jstadni2/pears_staff_report/tree/main/sample_outputs) directory.
