# utl-altair-personal-slc-comparing-two-excel-workbook-sheets
RE: Altair slc comparing two excel workbook sheets
    %let pgm=utl-altair-personal-slc-comparing-two-excel-workbook-sheets;

    %stop_submission;

    RE: Altair slc comparing two excel workbook sheets

    Too long to post here, see github

    github
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-comparing-two-excel-workbook-sheets


    community.altair
    https://community.altair.com/discussion/24939

    Compare two workbook sheet=ONE one with sheet=TWO
    ==================================================

                              d:/xls/onetwo.xlsx
                              ------------------
                       Sheet=ONE              Sheet=TWO

    ----------------------+                  -----------------------+
    | A1| fx    |STUDENT  |                  | A1| fx    |STUDENT   |
    ---------------------- --------------+   -----------------------------------------------
    [_] |   A   |  B |  C  |   D  |    E |   [_] |    A  | B  |   C |  D   |  E   |   F    |
    ---------------------- --------------|   ----------------------------------------------|
     1  |STUDENT|YEAR|STATE|GRADE1|GRADE2|    1  |STUDENT|YEAR|STATE|GRADE1|GRADE2|MAJOR   |
     -- |-------+----+-----+------+------|    -- |-------+---+------+------+------+--------|
     2  |JACK   |2020| NC  | 85   | 87   |    2  |JACK   |2020| NC  | 85   | 87   |Math    |
     -- |-------+----+-----+------+------|    -- |-------+----+-----+------+------+--------|
     3  |ALEX   |2025| MS  | 91   | 92   |    3  |ALEX   |2025| MS  | 90   | 92   |AR/VR   |
     -- |-------+----+-----+------+------|    -- |-------+----+-----+------+------+--------|
     4  |BARB   |2018| TN  | 78   | 92   |    4  |BARB   |2018| TN  | 78   | 92   |History |
     -- |-------+----+-----+------+------|    -- |-------+----+-----+------+------+--------|
     5  |MARY   |2020| NY  | 87   | 95   |    5  |MARY   |2020| MA  | 87   | 94   |Music   |
     -- |-------+----+-----+------+------|    -- |-------+----+-----+------+------+--------|
     6  |JEFF   |2025| NC  | 96   | 98   |    6  |JEFF   |2030| NY  | 96   | 98   |AI/ML   |
     -- |-------+----+-----+------+------|    -- |-------+----+-----+------+------+--------|
    [ONE]                                     7  |BILL   |2025| NC  | 82   | 96   |Robotics|
                                              -- |-------+----+-----+------+------+--------|
                                             [TWO]
    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|

     CREATE SHEETS
    */

    %utlfkil(d:/xls/onetwo.xlsx)

    libname xls excel "d:/xls/onetwo.xlsx";

    /*--- CREATE FIRST SAMPLE DATASET                ---*/

    data xls.one;
      input student$ year $ state $ grade1 grade2;
      label year = "Year of Birth";
      datalines;
    ALEX 2025 MS 91 92
    BARB 2018 TN 78 92
    JACK 2020 NC 85 87
    JEFF 2025 NC 96 98
    MARY 2020 NY 87 95
    ;
    run;

    /*--- CREATE SECOND SAMPLE DATASET               ---*/

    data xls.two;
      input student$ year $ state $ grade1 grade2 major $;
      label state = "Home State";
      datalines;
    ALEX 2025 MS 90 92 AR/VR
    BARB 2018 TN 78 92 History
    BILL 2025 NC 82 96 Robotics
    JACK 2020 NC 85 87 Math
    JEFF 2030 NY 96 98 AI/ML
    MARY 2020 MA 87 94 Music
    ;
    run;

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    /*--- RUN PROC COMPARE TO COMPARE THE TWO AHEETS ---*/

    &_init_;
    proc compare base=xls.'one$'n compare=xls.'two$'n;
      title "Comparison of sheet one with sheet two" listall;
      id student;
    run;

    libname xls clear;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    Comparison of sheet one with sheet two listall

    The COMPARE Procedure
    Comparison of XLS.one$ with XLS.two$
    (Method=EXACT)

                Data set summary

    Dataset            Nvars            Nobs
    ________________________________________
    XLS.one$               5    .
    XLS.two$               6    .

                        Variables summary

    Number of common variables:                             5
    Number of variables in XLS.two$ but not in XLS.one$:    1
    Number of ID statement variables:                       1

                           Observation summary

    Observation                  Base         Compare    ID
    _________________________________________________________________
    First Observation               1               1    STUDENT=ALEX
    First Unequal                   1               1    STUDENT=ALEX
    Last Unequal                    5               6    STUDENT=MARY
    Last Match                      5               6    STUDENT=MARY
    Last Observation                5               6    STUDENT=MARY


    Number of Observations in Common                               5
    Number of Observations in XLS.two$ but not in XLS.one$         1
    Total Number of Observations Read from XLS.one$                5
    Total Number of Observations Read from XLS.two$                6
    Number of Observations with Some Compared Variables Unequal    3
    Number of Observations with All Compared Variables Equal:      2


    Number of variables compared with all observations equal:                0
    Number of variables compared with some observations unequal:             4
    Total number of values which compare unequal:                            5
    Total number of values which compare equal:                             15
    Maximum Difference:                                                      1

                     Variables with unequal values

    Variable    Type          Length     Number diff        Max diff
    ________________________________________________________________
    YEAR        CHAR               8               1
    STATE       CHAR               4               2
    GRADE1      NUM                8               1               1
    GRADE2      NUM                8               1               1

             Value comparison results

                Base Value       Compare Value
    student     YEAR             YEAR
    __________________________________________
    JEFF        2025             2030

    Comparison of sheet one with sheet two listall

    The COMPARE Procedure
    Comparison of XLS.one$ with XLS.two$
    (Method=EXACT)

             Value comparison results

                Base Value       Compare Value
    student     STATE            STATE
    __________________________________________
    JEFF        NC               NY
    MARY        NY               MA

                             Value comparison results

                Base Value       Compare Value
    student            GRADE1           GRADE1           Diff.         % Diff.
    __________________________________________________________________________
    ALEX                   91               90              -1    -1.098901099

                             Value comparison results

                Base Value       Compare Value
    student            GRADE2           GRADE2           Diff.         % Diff.
    __________________________________________________________________________
    MARY                   95               94              -1    -1.052631579

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
