# UET Calendar

UET Calendar is a Google Sheet file with Apps Script program for making UET's school-schedule on Google Calendar.

## Installation

Make a copy of [this](https://docs.google.com/spreadsheets/d/19MJjkbqBNYJMGRkgw_0SipSdlCRQPTqaTgT69Ux-qtk/edit?usp=sharing) Google Sheet file.

## Usage

Info Sheet: 
- Change B1 to your calendar ID (ex: c_v7ss711irglpn681ua1sr5tado@group.calendar.google.com) or leave it blank (create new calendar)
- Change B2 to the start date of the school year (format Date time)
- Leave B3 as blank (auto fill with start date + 15*7 days)
<img src="https://user-images.githubusercontent.com/24197774/188897231-358993ec-3148-4db0-836b-2041fbe79c71.png" width="800">
<br>

UET Sheet:  
- Fill "Đăng ký" column with any value (ex: 1, a, ...)
<img src="https://user-images.githubusercontent.com/24197774/188651821-832e35d7-ad03-401a-ad6d-da3a44eaf1eb.png" width="800">

- Click "Code" then "Make Calendar" to export your calendar to Google Calendar
<img src="https://user-images.githubusercontent.com/24197774/188651988-eee62622-dfb3-4ced-aec9-41a71d937568.png" width="800">
* If it's the first time, Google will require permission to run, click Continue and Allow, click "Make Calendar" again to run
<br><br>

- Wait until the script is finnished then check your Google Calendar
<img src="https://user-images.githubusercontent.com/24197774/188652188-fafe6f3c-c802-4a27-a606-25cfdee95ad1.png" width="800">

- Other features:
  + "Add week number": add titles "Tuần 1, 2, 3, ..." to your calendar
  + "Delete Calendar": CAUTION: this will delete all of the data in your selected calendar
  + "Clear Sheet": clear selection in the "Đăng ký" column
- You can also add/remove subjects by making changes to "Đăng ký" column, click "Make Calendar" to update Google Calendar
- You can add new subjects, do sorting, filtering, . . . on the table, just make sure to reserve the correct form

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://choosealicense.com/licenses/mit/)
