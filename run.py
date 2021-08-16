from google.oauth2.service_account import Credentials
from datetime import datetime
from datetime import timedelta

import gspread
import random
import copy
import time
import smtplib

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('BIKES')

# Get data from google sheets
bikes_list = SHEET.worksheet('bike_list').get_all_values()
responses_list = SHEET.worksheet('form_responses').get_all_values()
sort_data = SHEET.worksheet('sort_data').get_all_values()
bookings_list = SHEET.worksheet('bookings')
update_bookings_list = SHEET.worksheet('bookings').get_all_values()
calendar2 = SHEET.worksheet('calendar2').get_all_values()
update_calendar2 = SHEET.worksheet('calendar2')
gs_size_guide = SHEET.worksheet('size_guide').get_all_values()

# Global Variables
booked_bikes = []
not_booked_bikes = []
bikes_dictionary = []
unavailable_bikes = []
bike_costs = []
total_cost = ""
hire_dates_requested = []
dates_filled_in_previous = sort_data[1][1]
sender = "bike_shop_owner@gmail.com"
receiver = responses_list[-1][3]
iterations = []

def error_func(this_error, error_comment):
    """
    This function is called up whenever there is an error
    It does not allow the booking to go ahead
    """
    print(f"The error is: {this_error}.")
    print(error_comment)
    print("The process will stop.")

    


def get_latest_response():
    """
    Fetch the latest form response from googlesheets and determine
    driving factors for bike selection
    """
    # get type and quantity of bikes selected
    # get heights
    # get date and duration of hire

    global types_list
    global heights_list

    # get the most recent form response from google sheets
    last_response = responses_list[-1]

    types_list_orig = [(last_response[7]),
                       (last_response[9]),
                       (last_response[11]),
                       (last_response[13]),
                       (last_response[15])]

    heights_list_orig = [(last_response[8]),
                         (last_response[10]),
                         (last_response[12]),
                         (last_response[14]),
                         (last_response[16])]

    global booking_number
    booking_number = len(responses_list)-1

    # max bikes hired = 5, but remove empty values from list
    types_list = list(filter(None, types_list_orig))
    heights_list = list(filter(None, heights_list_orig))

    # create dictionaries to store info about each bike requested
    for j in range(len(types_list)):
        d = {
            'booking_number': booking_number,
            'dates_of_hire': [],
            'bike_type': types_list[j],
            'user_height': heights_list[j],
            'bike_size': "",
            'possible_matches': [],
            'num_bikes_available': "",
            'price_per_day': "",
            'status': "Not booked",
            'comments': "",
            'booked_bike': "",
            'booked_bike_brand': "",
            'booked_bike_desc': "",
        }
        bikes_dictionary.append(d)

    global num_req_bikes
    num_req_bikes = len(bikes_dictionary)

    print(f"Booking number: {booking_number}")
    print(f"Submitted: {last_response[0]}")

    booking_processed(bikes_dictionary)


def booking_processed(bikes_dictionary):
    """
    This function determines if this booking number
    has already been run through the booking process
    If it has, stop
    """

    # for all bookings submitted
    for j in range(len(update_bookings_list)):

        # if the booking no. booked is same as current being processed
        if update_bookings_list[j][0] == str(booking_number):

            # print to terminal and stop process
            this_error = ("This booking has already been completed")
            error_comment = f"This booking was processed on {update_bookings_list[j][8]}"
            error_func(this_error, error_comment)

    match_size(bikes_dictionary)


def match_size(bikes_dictionary):
    """
    Match the heights specified in user form to the correct bike size
    """

    # iterate through list of bikes dictionaries (max 5)
    for i in range(len(bikes_dictionary)):

        # iterate through available bike sizes in size_guide google sheet
        for j in range(3, 9):

            # if the bike size in size_guide is the same as the user height
            # index the relevant bike size
            # and append to the dictionary
            if (gs_size_guide[j][4]) == bikes_dictionary[i]['user_height']:
                bikes_dictionary[i]['bike_size'] = gs_size_guide[j][8]

    match_price(bikes_dictionary)


def match_price(bikes_dictionary):
    """
    Fetch the price per day based on the bike type selected
    and append to dictionary
    """

    # iterate through list of bikes dictionaries (max 5)
    for i in range(len(bikes_dictionary)):

        # iterate through gs_bikes_list
        for j in range(len(bikes_list)):

            # if the bike type in the gs_bikes_list is same as that of the
            # dictionary, index the relevant price
            # and append to the dictionary
            if (bikes_list[j][4]) == bikes_dictionary[i]['bike_type']:
                bikes_dictionary[i]['price_per_day'] = bikes_list[j][5]

    find_unavailable_bikes(bikes_dictionary)


def find_unavailable_bikes(bikes_dictionary):
    """
    Define a list of unavailable bikes for those dates
    """

    global failed_booking

    # get users selected date
    users_start_date = responses_list[-1][5]
    today = datetime.now().date()

    # put date into different format
    start_date = datetime.strptime(users_start_date, "%m/%d/%Y").date()

    # calculate end date based on the hire duration
    delta = timedelta(days=int(responses_list[-1][6]) - 1)
    delta_day = timedelta(days=int(1))
    end_date = start_date + delta

    # list all dates in between and append to
    # hire_dates_requested
    # only do this ONCE
    if len(hire_dates_requested) == 0:
        while start_date <= end_date:
            hire_dates_requested.append(start_date.strftime("%m/%d/%Y"))
            start_date += delta_day

    # check dates submitted are not in the past
    if start_date < today:

        # if they are, process an email to inform user
        # of failed booking
        check_double_bookings()
        failed_booking = "Date of hire is in the past"
        return

    # for each date in the requested hire period
    for p in range(len(hire_dates_requested)):

        # and for each date in the calendar
        for x in range(len(calendar2[2])):

            # if the date is the one of those in the requested hire period
            if calendar2[2][x] == hire_dates_requested[p]:

                # for each bike index
                for q in range(len(calendar2)):

                    # if the cell is not blank, ie has a booking
                    # (do not count first 3 rows, and do not duplicate)
                    if calendar2[q][x] != "" and q > 3 and \
                            calendar2[q][0] not in unavailable_bikes:

                        # reference this bike and append to unavailable bikes
                        unavailable_bikes.append(calendar2[q][0])

    # also look for blanket unavailability in bikes list column G
    # only do this ONCE
    if len(iterations) == 0:

        # for each bike in bikes list
        for q in range(len(bikes_list)):

            # if Available? is No, append bike index to unavailable_bikes
            if bikes_list[q][6] == "No" and \
                    bikes_list[q][0] not in unavailable_bikes:

                unavailable_bikes.append(bikes_list[q][0])

    print(f">> unavailable bikes:- {unavailable_bikes}")
    match_suitable_bikes(bikes_dictionary)


def match_suitable_bikes(bikes_dictionary):
    """
    Use submitted form info to find selection of appropriate bikes
    """

    # loop through bikes requested, and compare to
    # bikes available in hire fleet
    # output the bike index of bikes which
    #  match type of those requested to the bike dictionaries
    for j in range(len(bikes_dictionary)):

        if bikes_dictionary[j]['status'] != "Booked":

            for i in range(len(bikes_list)):
                if bikes_dictionary[j]['bike_type'] == bikes_list[i][4] and \
                        bikes_dictionary[j]['bike_size'] == bikes_list[i][3]:

                    bikes_dictionary[j]['possible_matches'].\
                        append(bikes_list[i][0])

            bikes_dictionary[j]['num_bikes_available'] \
                = (len(bikes_dictionary[j]['possible_matches']))

    remove_unavailable_bikes(bikes_dictionary)
    book_bikes(bikes_dictionary)


def remove_unavailable_bikes(bikes_dictionary):
    """
    Cross reference possible matches in bike dictionaries with bike
    indexes in 'unavailable_bikes' list to check availability
    If not available, remove this bike index from the bike dictionary
    """

    # for each bike dictionary
    for j in range(len(bikes_dictionary)):

        # # only considering bikes who aren't already booked
        # if bikes_dictionary[j]["status"] != "Booked":

        # and for bike index in unavailable bikes
        for k in range(len(unavailable_bikes)):

            # if any of the bike indexes in unavailable bikes are found
            # in the bike dictionaries
            if unavailable_bikes[k] in bikes_dictionary[j]['possible_matches']:

                # then remove this bike index from the bike
                # dictionaries as it is not available for hire on
                # that date
                (bikes_dictionary[j]['possible_matches']).\
                    remove(unavailable_bikes[k])

    print(">> checking availability..")


def book_bikes_to_calendar(choose_bike_index):
    """
    Write the requested hire dates against the relevant bike index
    in google sheets
    """

    # for all bike indexes
    for x in range(len(calendar2)):

        # and for all dates listed in the calendar
        for z in range(len(calendar2[0])):

            # and for all hire dates requested
            for w in range(len(hire_dates_requested)):

                # match up the chosen bike index against the date
                # ensure cell is blank and we are not overwriting another date
                if (calendar2[x][0]) == str(choose_bike_index) and \
                        calendar2[2][z] == hire_dates_requested[w] and \
                        calendar2[x][z] == "":

                    # and write the booking number into that cell
                    update_calendar2.\
                        update_cell(x+1,
                                    z+1,
                                    bikes_dictionary[0]['booking_number'])

    print(f"Bike index {choose_bike_index} booked to calendar for {hire_dates_requested[0]} - {hire_dates_requested[-1]}")
    time.sleep(3)


def book_bikes(bikes_dictionary):
    """
    Determine how many matches in the 'possible matches'
    and call up book_bikes_to_calendar
    """

    global choose_bike_index
    choose_bike_index = ""

    # for each bike dictionary
    for j in range(len(bikes_dictionary)):

        if bikes_dictionary[j]['status'] != "Booked":

            # if the possible matches list is empty, move onto next j value
            # also add a comment
            if len(bikes_dictionary[j]['possible_matches']) == 0:

                bikes_dictionary[j]['comments'] = "No bikes available"
                continue

            # if the possible matches list = 1, then there is only 1 choice
            # so select that, call up book_bikes_to_calendar,
            # and remove it from other bike dictionaries 'possible matches'
            elif len(bikes_dictionary[j]['possible_matches']) == 1:

                # assign chosen_bike_index to this solo bike index
                choose_bike_index = bikes_dictionary[j]['possible_matches'][0]

                # call up book_bikes_to_calendar function
                book_bikes_to_calendar(choose_bike_index)

                # change status, update this bike dictionary, add bike
                # index to unavailable_bikes list, add this bike dict to
                # booked_bikes list
                bikes_dictionary[j]['status'] = "Booked"
                bikes_dictionary[j]['comments'] = ""
                bikes_dictionary[j]['booked_bike'] = choose_bike_index
                bikes_dictionary[j]['dates_of_hire'] = hire_dates_requested
                booked_bikes.append(bikes_dictionary[j])
                unavailable_bikes.append(choose_bike_index)

                if bikes_dictionary[j] in not_booked_bikes:
                    not_booked_bikes.remove(bikes_dictionary[j])

                # re-run remove_unavailable_bikes to remove this
                # bike index from other bike dicts
                remove_unavailable_bikes(bikes_dictionary)

            # if there is more than 1 bike available
            # randomly select a bike index from possible matches
            else:

                # randomly select one of the possible matches
                choose_bike_index = \
                    random.choice(bikes_dictionary[j]['possible_matches'])

                # run function with this bike index
                book_bikes_to_calendar(choose_bike_index)

                # change status, update this bike dictionary, add bike
                # index to unavailable_bikes list, add this bike dict to
                # booked_bikes list
                bikes_dictionary[j]['status'] = "Booked"
                bikes_dictionary[j]['comments'] = ""
                bikes_dictionary[j]['booked_bike'] = choose_bike_index
                bikes_dictionary[j]['dates_of_hire'] = hire_dates_requested
                unavailable_bikes.append(choose_bike_index)
                booked_bikes.append(bikes_dictionary[j])

                if bikes_dictionary[j] in not_booked_bikes:
                    not_booked_bikes.remove(bikes_dictionary[j])

                # re-run remove_unavailable_bikes to remove this
                # bike index from other bike dicts
                remove_unavailable_bikes(bikes_dictionary)

        continue

    booked_or_not(bikes_dictionary)


def find_alternatives(bikes_dictionary):
    """
    If the user is happy with alternatives, change the bike
    type (keep size and hire dates the same) and re-iterate
    with this new bike type
    """

    # only need to look for alternatives if there
    # are still bikes that aren't booked
    if not_booked_bikes != 0:

        global alt_bikes

        # for all bike dictionaries (now reduced in size from booked_or_not)
        for j in range(len(bikes_dictionary)):

            # list of all available bike types
            # keep inside j loop to restart from full list each time
            alt_bikes = ['Full suspension',
                         'Full suspension carbon',
                         'Full suspension carbon e-bike',
                         'Full suspension e-bike',
                         'Hardtail',
                         'Hardtail e-bike']

            # remove current bike type from alt_bikes so
            # function does not randomly choose the same bike type
            alt_bikes.remove(bikes_dictionary[j]['bike_type'])

            # choose random bike type (same size)
            bikes_dictionary[j]['bike_type'] = random.choice(alt_bikes)
            bikes_dictionary[j]['comments'] = "Finding alternative bike"

        # return to relevant function to perform again
        # only need to re-match the price, not the size as we know
        # the size is the same
        match_price(bikes_dictionary)


def booked_or_not(bikes_dictionary):
    """
    This function checks what the status is of all bike dictionaries
    If there are non-booked bikes and user is happy with alternative
    call up find_alternatives
    """
    # print(f"Number of iterations = {len(iterations)+1}")
    # time.sleep(5)

    # for all bikes dictionaries
    for j in range(len(bikes_dictionary)):

        # if the status does not equal Booked, they are
        # NOT booked, so append them to not_booked_bikes list
        if bikes_dictionary[j]['status'] != "Booked" and \
                 bikes_dictionary[j] not in not_booked_bikes:

            not_booked_bikes.append(bikes_dictionary[j])

        # if the bike has been booked but is still in not_booked_bikes
        # remove it
        if bikes_dictionary[j]['status'] == "Booked"\
                and bikes_dictionary[j] in not_booked_bikes:

            not_booked_bikes.remove(bikes_dictionary[j])

    # if all bikes have been booked
    if len(booked_bikes) == num_req_bikes:
        print("All bikes found.. sending confirmation emails")
        check_double_bookings()

    # if not all bikes have been booked
    # re-assign the bike dictionary to equal not_booked_bikes
    # to perform the iteration again for only these bikes
    elif responses_list[-1][17] == "Yes":
        bikes_dictionary = copy.copy(not_booked_bikes)
        iterations.append("1")

        # only allow max 4 iterations
        if len(iterations) > 4:

            # if reached max but bikes have been booked
            # send normal booking email
            if len(booked_bikes) > 0:
                check_double_bookings()

            # if nothing has been booked, raise error
            else:
                this_error = "Max iterations exceeded"
                failed_booking = "We could not find any suitable bikes"
                check_double_bookings()
                return

        find_alternatives(bikes_dictionary)

    # if the user does not want us to look for alternatives
    else:
        print(f"Bikes found:  {len(booked_bikes)}")
        print(f"Bikes not found:  {len(not_booked_bikes)}")
        print("User does not want alternatives")

        # but if some have been booked
        # ensure there are no double bookings
        check_double_bookings()


def check_double_bookings():
    """
    Perform a couple of calculations to ensure there
    are no double bookings
    """

    # FIRST CHECK
    # re-collect updated number from spreadsheet
    # (number counts num of cells filled in in calendar)
    sort_data2 = SHEET.worksheet('sort_data').get_all_values()
    dates_filled_in_now = sort_data2[1][1]

    # count how many dates have been filled in for this form
    num_dates_calendar = int(dates_filled_in_now) - \
        int(dates_filled_in_previous)

    # count how many dates should have been filled in
    num_dates_booked = len(booked_bikes)*len(hire_dates_requested)

    # check that these two numbers match
    if num_dates_booked != num_dates_calendar:
        error_comment = f"{num_dates_booked} should have been added to the calendar.\
            {num_dates_calendar} were added. \
            Please check any bikes/dates added to booking number {booking_number}."
        this_error = "The number of dates added to calendar is incorrect."
        error_func(this_error, error_comment)
        # send email to owner???

    print(">> checked calendar for double bookings")
    calculate_cost()


def calculate_cost():
    """
    This functions calculates the total cost of
    bike hire
    """
    global total_cost

    # for all bikes in booked list
    for i in range(len(booked_bikes)):

        # multiply price per day * number of hire days
        bike_costs.append(int(booked_bikes[i]['price_per_day']) * len(hire_dates_requested))

    total_cost = sum(bike_costs)

    bike_user_details()


def bike_user_details():
    """
    Now we know details of the booked bikes,
    collect last information about chosen bikes
    """

    # iterate through list of bikes dictionaries
    for i in range(len(booked_bikes)):

        # iterate through gs_bikes_list
        for j in range(len(bikes_list)):

            # if the booked bike index in the gs_bikes_list is same as that of
            # the dictionary, index the relevant brand and descr
            # and append to the dictionary
            if (bikes_list[j][0]) == booked_bikes[i]['booked_bike']:
                booked_bikes[i]['booked_bike_brand'] = bikes_list[j][1]
                booked_bikes[i]['booked_bike_desc'] = bikes_list[j][2]

    booking_details()


def booking_details():
    """
    This function organises the text to be sent
    in an email to user and owner to confirm booking
    """
    global user_email_subject

    # create string for hire dates
    if len(hire_dates_requested) > 1:
        user_email_subject = \
            f"{hire_dates_requested[0]} - {hire_dates_requested[-1]}"
    else:
        user_email_subject = {hire_dates_requested[0]}

    # create strings to append to emails
    # to detail booked and not booked bikes
    global email_booked_bike
    global email_not_booked_bike
    email_booked_bike = ""
    email_not_booked_bike = ""
    bike_throwaway = ""

    # for all booked bikes
    for i in range(len(booked_bikes)):

        # layout string of relevant details to be
        # added to email
        bike_throwaway = f"Bike {i+1}\n" \
                  f"Brand:          {booked_bikes[i]['booked_bike_brand']}\n" \
                  f"Description:    {booked_bikes[i]['booked_bike_desc']}\n" \
                  f"Type:           {booked_bikes[i]['bike_type']}\n" \
                  f"Rider height:   {booked_bikes[i]['user_height']}\n" \
                  f"Bike size:      {booked_bikes[i]['bike_size']}\n" \
                  f"Price per day:  {booked_bikes[i]['price_per_day']} GBP \n" \
                  f"Total price for {len(hire_dates_requested)} days = {bike_costs[i]} GBP\n" \
                  '\n'

        email_booked_bike += bike_throwaway

    # for all not booked bikes
    for i in range(len(not_booked_bikes)):

        # layout string of relevant details to be
        # added to email
        bike_throwaway = f"Type:           {not_booked_bikes[i]['bike_type']}\n" \
                f"Rider height:   {not_booked_bikes[i]['user_height']}\n" \
                f"Bike size:      {not_booked_bikes[i]['bike_size']}\n" \
                    '\n'

        email_not_booked_bike += bike_throwaway

    if len(booked_bikes) == 0:
        email_booked_bike = "None"
    if len(not_booked_bikes) == 0:
        email_not_booked_bike = "None"

    # only write booking to booking list if bikes have been booked
    if len(booked_bikes) > 0:
        add_booking_to_gs()


def add_booking_to_gs():

    # find last row in bookings list
    last_row_in_bookings_list = len(update_bookings_list) + 1

    global booked_bikes_list
    booked_bikes_list = ""

    # for all bikes that have been booked
    for i in range(len(booked_bikes)):

        # append bike index to a string
        booked_bikes_list += (booked_bikes[i]['booked_bike']) + ', '

    # update columns in gs bookings 
    bookings_list.update_cell(last_row_in_bookings_list, 1, booking_number)
    bookings_list.update_cell(last_row_in_bookings_list, 2, responses_list[-1][0])  # noqa
    bookings_list.update_cell(last_row_in_bookings_list, 3, hire_dates_requested[0])  # noqa
    bookings_list.update_cell(last_row_in_bookings_list, 4, hire_dates_requested[-1])  # noqa
    bookings_list.update_cell(last_row_in_bookings_list, 5, len(booked_bikes))
    bookings_list.update_cell(last_row_in_bookings_list, 6, responses_list[-1][1])  # noqa
    bookings_list.update_cell(last_row_in_bookings_list, 7, responses_list[-1][2])  # noqa
    bookings_list.update_cell(last_row_in_bookings_list, 8, booked_bikes_list)
    bookings_list.update_cell(last_row_in_bookings_list, 9, str(datetime.now()))  # noqa

    print(">> booking added to bookings list in Google Sheets")


get_latest_response()


# create strings for emails (this doesn't work in a function...)
if len(booked_bikes) > 0:

    subject = f"Bike hire booking confirmed {user_email_subject}"

    message_to_user = f"""\
Subject: {subject}
To: {receiver}
From: {sender}

Hello {responses_list[-1][1].split(' ', 1)[0]}!

Thank you for booking with us.  Below are your booking details:
Booking name:               {responses_list[-1][1]}
Booking contact number:     {responses_list[-1][2]}
Dates of hire:              {user_email_subject}
Number of bikes booked:     {len(booked_bikes)}

Booked bikes:
{email_booked_bike}

Bikes we could not book:
{email_not_booked_bike}

Important Information:
Time out:  9am on first day of hire
Time due back:  4.45pm on last day of hire

Terms of Hire:
You may cancel your booking any time up until 24 hrs
before the first date of hire.
The total payable amount must be paid before you turn up,
or on the first date of hire.

Please read the safety brief provided on the website.
If you are due to be late back, please let the shop owner know.
There may be additional charges for returning after the time due back.

Payment:
You can pay by phone on 08796 236458, or pay on the
first day of hire by card or cash.
The total amount payable for this booking is: {total_cost} GBP

Want to amend your booking?
Please call 08796 236458 or email info@bike_shop.com

See you soon!

Regards
Bike Shop
"""

    message_to_owner = f"""\
Subject: {subject}
To: {sender}
From: {sender}

The following bikes have been booked for {user_email_subject}:

Booking name:               {responses_list[-1][1]}
Booking contact number:     {responses_list[-1][2]}
Dates of hire:              {user_email_subject}
Number of bikes booked:     {len(booked_bikes)}
Total cost:                 {total_cost}

Booked bikes:
{email_booked_bike}

Bikes we could not book:
{email_not_booked_bike}

End of email
"""

else:

    subject = f"Sorry - we were not successful with your booking for {user_email_subject}" # noqa

    message_to_user = f"""\
Subject: {subject}
To: {receiver}
From: {sender}

Hello {responses_list[-1][1].split(' ', 1)[0]}

Unfortunately we could not complete your booking.

If you'd like to enquire further, please phone us on
08796 236458 or email us at bike_shop_owner@gmail.com
and we'd be happy to help you arrange something else.

Regards
Bike Shop
        """

    message_to_owner = f"""\
Subject: Hire form submitted but not processed
To: {sender}
From: {sender}

There was a booking form submitted by {responses_list[-1][1]}
at {responses_list[-1][0]}, but the booking could not be processed because:

{failed_booking}.

Consult the Google Sheets database for information about the booking.
        """

# send the emails
with smtplib.SMTP("smtp.mailtrap.io", 2525) as server:
    server.login("a3b48bd04430b7", "0ec73c699a910c")
    server.sendmail(sender, receiver, message_to_user)
    server.sendmail(sender, receiver, message_to_owner)


print(">> END OF SCRIPT")
raise SystemExit
