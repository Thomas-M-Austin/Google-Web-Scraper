import requests
import json
import pandas as pd
from datetime import date
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QStatusBar, QWidget

class SearchWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # create central widget and layout
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.layout = self.central_widget.layout()

        # create api key label and input field
        self.api_key_label = QLabel('API key:', self.central_widget)
        self.api_key_label.move(50, 25)
        self.api_key_input = QLineEdit(self.central_widget)
        self.api_key_input.move(150, 25)

        # create site name label and input field
        self.site_name_label = QLabel('Site name:', self.central_widget)
        self.site_name_label.move(50, 75)
        self.site_name_input = QLineEdit(self.central_widget)
        self.site_name_input.move(150, 75)

        # create keyword label and input field
        self.keyword_label = QLabel('Keyword:', self.central_widget)
        self.keyword_label.move(50, 125)
        self.keyword_input = QLineEdit(self.central_widget)
        self.keyword_input.move(150, 125)

        # create search button
        self.search_button = QPushButton('Search', self.central_widget)
        self.search_button.move(150, 175)
        self.search_button.clicked.connect(self.search)

        # create status bar
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

        # set window title, size and show the window
        self.setWindowTitle('Scale Serp Searcher')
        self.setGeometry(100, 100, 400, 250)
        self.show()

    def search(self):
        # get values from input fields
        api_key = self.api_key_input.text()
        site_name = self.site_name_input.text()
        keyword = self.keyword_input.text()

        # check if fields are empty and show error messages
        if not api_key:
            self.status_bar.showMessage('Please enter API key')
            return
        if not site_name:
            self.status_bar.showMessage('Please enter site name')
            return
        if not keyword:
            self.status_bar.showMessage('Please enter keyword')
            return

        # set up the request parameters
        params = {
            'api_key': api_key,
            'q': f'site:{site_name} in-title:"{keyword}"',
            'location': 'United States',
            'google_domain': 'google.com',
            'gl': 'us',
            'hl': 'en',
            'time_period': 'last_year',
            'sort_by': 'date',
            'num': '100',
            'include_html': 'false',
            'output': 'json'
        }

        # update status bar and search button text
        self.status_bar.showMessage('Searching...')
        self.search_button.setText('Searching')

        # make the http GET request to Scale SERP
        api_result = requests.get('https://api.scaleserp.com/search', params)

        # load the JSON response into a pandas DataFrame
        df = pd.json_normalize(api_result.json()['organic_results'])

        # create output file name based on search query and today's date
        q_param_value = params['q'].replace(':', '_')
        output_file_name = f"{q_param_value}_{date.today().strftime('%Y-%m-%d')}.xlsx"

        # write the DataFrame to an Excel file with the output file name
        df.to_excel(output_file_name, index=False)

        # update status bar and show success message
        self.status_bar.showMessage('File successfully saved')
        
        # clear input fields
        self.site_name_input.clear()
        self.keyword_input.clear()
        self.search_button.setText('Search')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    search_window = SearchWindow()
    sys.exit(app.exec_())
