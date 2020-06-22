## About
This project is to automate COAP MTech offers.

## Pre-requisites
* Before running any script it is important to backup all Excel (.xlsx) files, to rollback on errors.
* Create an empty offers file called <PREFIX>_offers.xlsx with empty round sheets for all rounds with title “Round_1”, “Round_2”… etc.
* Create an empty summary file called <PREFIX>_summary.xlsx with empty round sheets for all rounds. Pre-fill in first round details with number of seats, multipliers, etc.

## Files
The code requires the following files:
* **Applicant Input File**: This is the master list of applicants as an Excel file. It is **absolutely important**
to arrange the data in the order mentioned below.

    A=coap_id, B=gate_score, C=appl_id, D=name=, E=gender, F=category, G=disabled_flg,
    H=gate_id, I=btech_score_1, J=btech_score_2, K=email, L=mobile, M=gate_stream, N=btech_stream

* **"PREFIX"_summary File**: This is a high-level summary file that contains high-level summary information per
    category for every round (in a sheet of its own). Namely, the number of seats remaining, the final cutoffs, multiplication factors (if you would like to make multiple offers per seat)

* **"PREFIX"_offers File**: This is the detailed offers file that contains details about the students who have been
    made offers to in each round (in a sheet of its own). Each row in this file will contain the following information.

    coap_id, status (IITH offer status [Accept | Retain | Reject]), reason (reason for offer status),
    name, gender, student_category, offer_seat_category, appl_id, gate_id, disabled_flag (Y | N),
    gate_score, btech_score, email, mobile, gate_stream, and btech_stream.

## Making Offers
Prior to making offers you must copy the summary details from the previous round's sheet and fill up this round's
correct summary details, i.e., how many seats left over per category and what the multipliers are.
Run **make_offers.py** when making an offer during a round. This program will load all previous round's offers and check
the candidate's offer status when making an offer in the current round. If a _candidate has rejected our offer in a previous
round_ or was _not made an offer by IITH and accepted another institute's offer_, then this candidate is not made an offer
this round. If a _candidate either retains or accepts our offer_, we assign an offer to this candidate.
Each category's cutoffs are also automatically computed and recorded in the **"PREFIX"_summary file**.

The help file for make_offers.py
```
$ python3 make_offers.py --help
usage: make_offers.py [-h] -a APPLICANTS_FILE -o OFFERS_PREFIX -r ROUND

optional arguments:
  -h, --help            show this help message and exit
  -a APPLICANTS_FILE, --applicants_file APPLICANTS_FILE
                        The master file containing all the applications with
                        coap_id, gate_id, appl_id
  -o OFFERS_PREFIX, --offers_prefix OFFERS_PREFIX
                        This prefix will use <prefix>_offers.xlsx and
                        <prefix>_summary.xlsx files.
  -r ROUND, --round ROUND
                        Current round of offers to make
```
**Sample usage:**
```
$ python3 make_offers.py --applicants_file "sample_app_file.xlsx" --offers_prefix "SAMPLE_TA" --round 1
```
Here “sample_app_file.xlsx” is the master file with the list of applications. Using the —offers_prefix “SAMPLE_TA” will look for two files: SAMPLE_TA_offers.xlsx and SAMPLE_TA_summary.xlsx. The first one is the file with all the students that are offered seats by us in Round 1 and the second is the overall summary of how many seats remain to be offered for each category.

## Update Offers
Run **update_offers.py** when making status updates on receiving COAP update files for latest round offers.
Here, we update the status of our latest round's offers depending on whether an applicant _makes a decision
about our offer (-our flag)_ or _decides to choose another institute's offer (-oth flag)_.

It is important to note that the program can only handle 3 update files and **they have to be processed in the given order**:
* **< Round X > IIT Hyderabad Candidate Decision File**

* **< Round X > IIT Hyderabad Offered But Accept and Freeze at Other File**

* **< Round X > Consolidated Accept and Freeze Candidates Across All Institutes File**

The help file for update_offers.py
```
$ python3 update_offers.py --help
usage: update_offers.py [-h] -a APPLICANTS_FILE -u UPDATE_FILE -c COAP_ID_COL
                        -op OFFERS_PREFIX [-prg PROGRAM] [-pcol PROGRAM_COL]
                        -r ROUND [-our OUR_STATUS_COL | -oth OTHER_STATUS_COL]

optional arguments:
  -h, --help            show this help message and exit
  -a APPLICANTS_FILE, --applicants_file APPLICANTS_FILE
                        The master file containing all the applications with
                        coap_id, gate_id, appl_id
  -u UPDATE_FILE, --update_file UPDATE_FILE
                        The coap updates file containing all the application
                        status changes
  -c COAP_ID_COL, --coap_id_col COAP_ID_COL
                        The column name in updates spreadsheet where coap_id
                        is present. E.g. A means look in column A
  -op OFFERS_PREFIX, --offers_prefix OFFERS_PREFIX
                        This prefix will use <prefix>_offers.xlsx and
                        <prefix>_summary.xlsx files.
  -prg PROGRAM, --program PROGRAM
                        This is the program offered (e.g. CSE, NIS)
  -pcol PROGRAM_COL, --program_col PROGRAM_COL
                        This is the column name in updates spreadsheet where
                        program offered is present.
  -r ROUND, --round ROUND
                        Current round of offers to update
  -our OUR_STATUS_COL, --our_status_col OUR_STATUS_COL
                        The column in updates spreadsheet where iith_status is
                        present.
  -oth OTHER_STATUS_COL, --other_status_col OTHER_STATUS_COL
                        The column in updates spreadsheet where other_status
                        is present.
```
**Sample usage:**
```
python3 update_offers.py -a "sample_app_file.xlsx" -u "round1_update1.xlsx" -c A -op "SAMPLE_TA" -r 1 -oth N
```
This will look at update file “round1_update1.xlsx” find COAP_ID in column A, offer_files (summary and detail) are prefixed with SAMPLE_TA, round 1 updates, and other_Status of file found in col N is processed.

## Sample Rounds
* Round 1 (with no prior offers is tested)
    `python3 make_offers.py -a "sample_app_file.xlsx" -o "SAMPLE_TA" -r 1`

* Round 1 update 1: with other offer taken, our reject
    `python3 update_offers.py -a "sample_app_file.xlsx" -u "round1_update1.xlsx" -c A -op "SAMPLE_TA" -r 1 -oth N`
* Round 1 update 2: with our offer retained, rejected
    `python3 update_offers.py -a "sample_app_file.xlsx" -u "round1_update2.xlsx" -c A -op "SAMPLE_TA" -r 1 -our J`

* Round 2 (with prir offers made)
    `python3 make_offers.py -a "sample_app_file.xlsx" -o "SAMPLE_TA" -r 2`



* Round 1 update 1: Round 1 IIT Hyderabad Candidate Decision Report.xlsx
```
python3 update_offers.py -a "sample_app_file.xlsx" -u "./coapround1decision/Round 1 IIT Hyderabad Candidate Decision Report.xlsx" -c A -op "SAMPLE_TA" -r 1 -our J -prg "CSE" -pcol H
```
This will check the round 1 offers made in SAMPLE_TA_* (summary and offers) files and compare against the status changes in IITH Candidate Decision Report. The COAP_ID in cand. decision report is located in column A (-c A) for round 1 (-r 1)
and our offer's status is in column J (-our J) and the program offered is "CSE" (you can modify this for your dept/stream) and the column H holds the name of the program in the file.

* Round 1 update 2: Round 1 IIT Hyderabad Offered But Accept and Freeze at Other....xlsx
```
python3 update_offers.py -a "sample_app_file.xlsx" -u "./coapround1decision/Round 1 IIT Hyderabad Offered But Accept and Freeze at Oth....xlsx" -c A -op "SAMPLE_TA" -r 1 -oth N -prg "CSE" -pcol H
```
This will check the round 1 offers made in SAMPLE_TA_* (summary and offers) files and compare against the status changes in IITH Offered But Accept and Freeze at Other Report. The COAP_ID in update report is located in column A (-c A) for round 1 (-r 1) and our other's status is in column N (-oth N) and the program offered is "CSE" (you can modify this for your dept/stream) and the column H holds the name of the program in the file.

* Round 1 update 3: Round 1 Consolidated Accept and Freeze Candidates Across All Institutes.xlsx
```
python3 update_offers.py -a "sample_app_file.xlsx" -u "./coapround1decision/Round 1 Consolidated Accept and Freeze Candidates Across All Institutes.xlsx" -c A -op "SAMPLE_TA" -r 1 -oth H
```
This will check the round 1 offers made in SAMPLE_TA_* (summary and offers) files and compare against the status changes in IITH Consolidated Accept and Freeze Candidates Across All Institutes Report. The COAP_ID in update report is located in column A (-c A) for round 1 (-r 1) and our other's status is in column H (-oth H). Note there is no program information in
this file.


## Minor Bugs and Workarounds
Do not manually edit files because it sometimes leaves “blank cells” at the bottom which then causes a bug in the code. If this happens, open the .xlsx file and manually delete rows under the rows which have content, till the bug doesn’t appear.

## Common Pitfalls
* At the very first make_offer run, not ensuring that you have summary and offers files with empty sheets per round, each
named "Round_X", where X is the round number.

* Not having checked the summary details before running make_offer for a new round. This file **MUST** contain the correct number of seats remaining and multipliers before the offers for the round can be made.

* Not backing up files after running each update. In case of an error, a backup comes very handy.

* Running update_offers on round_X but not setting flag "-r X" correctly. For example. it is possible to make the error of updating the round 2 offers with -r 1 (round 1) flag set. This will cause errors in your output files.
