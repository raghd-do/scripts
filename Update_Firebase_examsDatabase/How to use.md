# رفع يانات السجل الأكاديمي من قاعدة الإكسل إلى قاعدة Firebase

### First: convert excle to json
1. open `السجل الأكاديمي.xlsx` and insert the data that you want to be uplauded to the Firebase
2. run `excel to json.py` on the Termenal, to genetate a json object of the exele data
3. a `grouped_data.json` will be genarated

### Second: downlaud admin privet key file
1. go to the `Exams` Firebase console
2. open project sittings
3. open `Service accounts`
4. from `Admin SDK configuration snippet` choose `Python`
5. click `Generate new private key` button, to downlaud the admin privet key, then store it some where outside the project folder ***MAKE SURE NOT TO EXPOSET TO PUBLIC*** like `GitHub`

> NOTE: this file will allow you to modify the Firebase database from your local code as an admin

### Third: uplaud genarated student exams JSON file to Firebase using the `admin privet key` file

> I'm working on cleaning the code :P