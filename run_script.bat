:: Build the Docker image for pears_staff_report.py
docker build -t il_fcs/pears_staff_report:latest .
:: Create and start the Docker container 
docker run --name pears_staff_report il_fcs/pears_staff_report:latest
:: Copy /sample_outputs from the container to the build context
docker cp pears_staff_report:/pears_staff_report/sample_outputs/ ./
:: Remove the container
docker rm pears_staff_report
pause
