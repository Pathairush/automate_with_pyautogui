# SAP Report Extraction with Pyautogui

Hi everyone :wave: , welcome to `SAP Report Extraction with Pyautogui`. In this repository we provide a snippet code to extract the report from `ZFIARAGING` module throguh the SAP LOGON GUI. In short `ZFIARAGING` SAP module provide you a collection report for overdue account receivable (the report about how long overdue payment is).

 The report extraction is the most tedious task I have ever work with :fire: . You just have to pass the same parameter every day to get a new report. Also, you can't export the end result because all the calculation is processed on the fly when you query it. The process take almost an hour to finish. What a waste of time. :x:
 
 Thus, we aim to automate this process so that we can have time to do something that is more important.

## Required
1. SAP GUI Logon application
2. VPN Connection application (if needed)

## How to run this project

```
python run_sap_export.py
```

## Todos

- [X] Refactor the script to receive a `keyword argument` from terminal
- [X] Refactor the code to provide different input parameter for different `DEV`, and `PRD` servers.
- [X] Add kill process before running the sap report extaction.
- [ ] Schedule the task with `Apache airflow`
- [ ] Consolidate the vpn connection, report extraction, and upload to S3 together
- [ ] Provide mockup files for configuration and location reference

## Note

Because the config and location reference file has contained the confidentials information, I can't put it here in the reporsitory. I will provide you with the mockup template instead.

## Resources
- [Manually raising (throwing) an exception in Python](https://stackoverflow.com/questions/2052390/manually-raising-throwing-an-exception-in-python)