# رفع بيانات السجل الأكاديمي من قاعدة الإكسل إلى قاعدة Firebase

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
1. on the terminal, dounloud the dependensy library to work in `upluad-to-firebase.py` file like so
    ```bash
    pip install firebase_admin
    ```
2. open `upluad-to-firebase.py` file and change the `private_key` varible to your firebase credential file path
    ```python
    #Write the [File Path] to your firebase credential file
    private_key = "fir-web-codelab-79cbb-firebase-adminsdk-4ebbq-fc38a1d45d.json"
    ```
3. run the `upluad-to-firebase.py` file like so on the terminal
    ```bash
    python upluad-to-firebase.py
    ```
    > NOTE: remember to `cd` to directory of the file before running the command :)