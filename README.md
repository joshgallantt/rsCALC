<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/joshgallantt/rsCALC">
    <img src="assets/dress.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">rsCALC</h3>

  <p align="center">
    A data manager and calculator for RewardStyle influencers
    <br />
  </p>
</p>

<!-- ABOUT THE PROJECT -->
## About The Project

If you've tried to grab insights from RewardStyles' commission interface, you know that it's dated and unwieldy. rsCALC is designed to automate the process and enable you to make business decisions faster.

Features:
* Automatically logs in to your RewardStyle account and downloads your data from a specified date range
* Provides some at-a-glance stats for the given period:
  * Estimated Advertiser Earnings (using automatically updated brand rates)
  * Top Brands
  * Top Products
  * Top Refunds
* Combines and formats all of the data into a .csv file for further use
* (Untested) Besides Windows, rsCALC should work on pre M1 Macs, and Linux

### Built With

* [Python 3.X](https://www.python.org/downloads/)
* [Selenium](https://github.com/SeleniumHQ/Selenium)
* [Pandas](https://pandas.pydata.org/)
* [Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
* [Chromium/ChromeDriver](https://chromedriver.chromium.org/)
* [Pillow](https://python-pillow.org/)
* [Tabulate](https://github.com/astanin/python-tabulate)


<!-- GETTING STARTED -->
## Getting Started

1. Install Python 3.X and make sure it's defined to the Path on installation.
2. Have Google Chrome installed on your machine
3. Install the following python libraries:

  ```sh
  pip install selenium pandas Pillow bs4 lxlm
  ```
4. Extract the zip anywhere on your machine and run rsCALC.py

<!-- USAGE EXAMPLES -->
## Usage

1. Enter your RewardStyle login info
2. Click Start and wait for the files to be downloaded/worked with
3. After you see the output you can export the data for your range using the Export button
4. If you wish you see another date range, just select those dates and start again.

Note: Anything not exported gets deleted on program exit.

![](example.gif)


<!-- CONTRIBUTIONS -->
## Contributing

See the [open issues](https://github.com/joshgallantt/rsCALC/issues) for a list of proposed features (and known issues). I'm new to programming and if you have any experience with Python and any of the frameworks, please feel free to submit a pull request or DM me.

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.

<!-- CONTACT -->
## Contact

* Instagram - [@joshgallantt](https://instagram.com/joshgallantt)

* Project Link: [https://github.com/joshgallantt/rsCALC](https://github.com/joshgallantt/rsCALC)

* Discussions [here](https://github.com/joshgallantt/rsCALC/discussions)

<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements
* [Icon from Flaticon](https://www.flaticon.com/free-icon/dress_1785255?term=dress&page=1&position=2&page=1&position=2&related_id=1785255&origin=search)

* Thank you James for answering way too many of my coding questions!

* Not affiliated with RewardStyle, please read the license for more info.

