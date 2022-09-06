# UET Calendar

UET Calendar is a Google Sheet file with Apps Script program for making UET's school-schedule on Google Calendar.

## Installation

Make a copy of [this](https://docs.google.com/spreadsheets/d/19MJjkbqBNYJMGRkgw_0SipSdlCRQPTqaTgT69Ux-qtk/edit?usp=sharing) Google Sheet file.

## Usage

Info Sheet: 
- Change B1 to your calendar ID (ex: c_v7ss711irglpn681ua1sr5tado@group.calendar.google.com) or leave it blank (will create a new calendar)
- Change B2, B3 to the start, end date of the calendar (format Date time)
- B3 can be blank (auto fill with start date + 15 weeks)
<img src="https://user-images.githubusercontent.com/24197774/188657847-3e699b83-c7b8-4283-9ac1-f9bf1b36a0c4.png" width="800">
<br>

UET ... Sheet:  
- Find subjects of your choice and fill its "Đăng ký" column with any value (ex: 1, a, ...)
<img src="https://user-images.githubusercontent.com/24197774/188651821-832e35d7-ad03-401a-ad6d-da3a44eaf1eb.png" width="800">
* Only select school subject
<br><br>

- Click "Code" then "Make Calendar" to export your calendar to Google Calendar
<img src="https://user-images.githubusercontent.com/24197774/188651988-eee62622-dfb3-4ced-aec9-41a71d937568.png" width="800">
* If it's the first time, Google will require permission to run, click Continue, Your google account and Allow, click "Make Calendar" again
<br><br>

- Wait until the script is finnished then check your Google Calendar
<img src="https://user-images.githubusercontent.com/24197774/188652188-fafe6f3c-c802-4a27-a606-25cfdee95ad1.png" width="800">

- (Optional) Change color of your calendar (because default color bad :v)
- Other choices:
  + "Add week number": add "Tuần 1, 2, 3, ..." to your calendar
  + "Delete Calendar": CAUTION: this will delete all of the data in your selected calendar
  + "Clear Sheet": clear selection in the "Đăng ký" column
- You can also change your choices in Google Sheet then click "Make Calendar" to update calendar
- You can add new subjects, do some sorting, filtering, . . ., just make sure to reserve the correct form of the table

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://choosealicense.com/licenses/mit/)
