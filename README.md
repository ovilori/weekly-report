# weekly-report
This is a personal project including codes I wrote to solve a personal challenge at work, where I was responsible for the weekly uptime report of all service providers on our network. Preparing the report was time consuming for the following reasons:

- Log in to the solarwind on a web browser. 
- Navigate to the report section and manually edit the start and end date of the new week.
- Svae, and then export the data to an excel sheet.
- Manually clean the data.
- Manually create new sheets for all providers in excel.
- Manually copy the data for each provider to the seperate sheets for each days of the week, etc.

This took more time to complete and I was able to reduce the time it takes to prepare the report by about 80%. Now, the python code performs the data extraction and sorts accordingly. The code also take care of addition of new providers/locations in future.

## Weekly Provider Uptime Report 

Files: exportProviderUptime.py & sortProviderUptime.py

### Weekly Location LAN Uptime Report 

Files: exportLANUptime.py & sortLANUptime.py
