# Awork Parser

## Installation and usage
1. Install [python3](https://www.python.org/) for your operating system
2. `git clone` this repository
3. `cd` to the repository directory
4. Fill out the predictions.json file. dayOfWeek is in isoweekday format (Monday = 1, Sunday = 7)
5. `pip install -r requirements.txt`
6. Visit awork.io `/time-tracking/my-day` to download your Excel export
    ![awork](https://i.imgur.com/ekQzlcJ.png)
7. Put the Excel export in the `excel_sheets` directory
8. Run `python main.py`
9. Profit!

Note: Ignore the warning about default style I don't know how to suppress it.
