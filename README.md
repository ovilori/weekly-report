# weekly-report
This is a personal project including codes I wrote to solve a challenge at work, where I was responsible for the weekly uptime report of all service providers on our network. Preparing the report was time consuming for the following reasons:

- Log in to the solarwind on a web browser. 
- Navigate to the report section and manually edit the start and end date of the new week.
- Save, and then export the data to an excel sheet.
- Manually clean the data.
- Manually create new sheets for all providers in excel.
- Manually copy the data for each provider to the seperate sheets for each days of the week, etc.

Preparing the report took more days due to the manual work involved. I wrote Python script to generate, and email the network availability report for 15+ providers across 120+ sites, reducing turnaround time by about 95%. The code is also written to in such a way that onboarding new providers, or adding more sites in the future will not result in an error.

## Weekly Provider Uptime Report 

Files: exportProviderUptime.py & sortProviderUptime.py

### Weekly Location LAN Uptime Report 

Files: exportLANUptime.py & sortLANUptime.py
