FROM python:3.9

WORKDIR /pears_staff_report

COPY . .

RUN pip install -r requirements.txt

CMD [ "python", "./pears_staff_report.py" ]