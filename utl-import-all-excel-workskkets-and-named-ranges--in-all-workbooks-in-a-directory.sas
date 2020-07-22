Import all excel worksheets and named ranges  in all workbooks in a directory

github
https://tinyurl.com/yyzvop49
https://github.com/rogerjdeangelis/utl-import-all-excel-workskkets-and-named-ranges--in-all-workbooks-in-a-directory

SAS-L
https://listserv.uga.edu/cgi-bin/wa?A2=SAS-L;5a8acd9c.2007c

The difficult part is creating the input test case.

After that just

* SOLUTION;

options validmemname= extend;
libname xel ("d:/xel/ages.xlsx", "d:/xel/genders.xlsx");
proc copy in=xel out=work;
run;quit;

  Note
     a. dosubl can process each sehet and named range and provide checkpoint/restart  processing, error logging and
        rename 'sheet1$' to valid unquoted name.
        (solution not shown)
     b. do_over can provide renaming. Not as easy to do checkpoint/restart. (not shown)
     c. I feel 'call execute'and 'proc import/export' provide less flexible solution?
     d. proc contents gives you a quick list where you can decide what you want to import.
                 __       _             _                  _
 _   _ ___  ___ / _|_   _| |  ___ _ __ (_)_ __  _ __   ___| |_
| | | / __|/ _ \ |_| | | | | / __| '_ \| | '_ \| '_ \ / _ \ __|
| |_| \__ \  __/  _| |_| | | \__ \ | | | | |_) | |_) |  __/ |_
 \__,_|___/\___|_|  \__,_|_| |___/_| |_|_| .__/| .__/ \___|\__|
                                         |_|   |_|

* Snippet of code to get a list of workooks in a directory - not used in solution;
data workbooks;
  _nam=filename('fid',"&dir");
  _opn=dopen('fid');
  do _n_=1 to dnum(_opn);
     filename="&dir/"!!dread(_opn,_n_);
     output;
     drop _:;
  end;

run;quit;


WORK.WORKBOOKS total obs=2

Obs         filename

 1     D:/xel/ages.xlsx
 2     D:/xel/genders.xlsx

*_                   _
(_)_ __  _ __  _   _| |_
| | '_ \| '_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
;

* create folder for execl file;

data _null_;
  length rc $2;
  rc=dcreate("xel","d:/");
run;quit;

* create two workbooks with 2 sheets and two named ranges in each;


/* just in case there are assignments;
libname sex clear;
libname age clear;
libname xel clear;

* just in case they exist;
%utlfkil(d:/xel/genders.xlsx);
%utlfkil(d:/xel/ages.xlsx);


* create workbooks;
libname sex    "d:/xel/genders.xlsx";
libname age    "d:/xel/ages.xlsx";

data
      sex.males  (where=(sex="M"))
      sex.females(where=(sex="F"))
      age.agege13   (where=(age ge 13))
      age.agele13   (where=(age le 13))
 ;
     set sashelp.class;
run;quit;

*     _               _      _                   _
  ___| |__   ___  ___| | __ (_)_ __  _ __  _   _| |_
 / __| '_ \ / _ \/ __| |/ / | | '_ \| '_ \| | | | __|
| (__| | | |  __/ (__|   <  | | | | | |_) | |_| | |_
 \___|_| |_|\___|\___|_|\_\ |_|_| |_| .__/ \__,_|\__|
                                    |_|
;

libname xel ("d:/xel/ages.xlsx", "d:/xel/genders.xlsx");

ods output members=mbrs;
ods output directory=dirs;
proc contents data=xel._all_;
run;quit;

libname xel clear;

*            _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| '_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
;

WORK.DIRS total obs=8

                                                  n
Obs    Label1           cValue1                Value1

 1     Libref           XEL                       .
 2     Levels           2                         2
 3     Engine           EXCEL                     .
 4     Physical Name    d:/xel/ages.xlsx          .
 5     User             Admin                     .
 6     Engine           EXCEL                     .
 7     Physical Name    d:/xel/genders.xlsx       .
 8     User             Admin                     .

WORK.MBRS total obs=8

                          Mem
Obs    Num    Name        Type    Level    Obs    Vars    Label    DBMSTYPE

 1      1     agege13     DATA      1       .       5               TABLE
 2      2     agege13$    DATA      1       .       5               TABLE
 3      3     agele13     DATA      1       .       5               TABLE
 4      4     agele13$    DATA      1       .       5               TABLE

 5      5     females     DATA      2       .       5               TABLE
 6      6     females$    DATA      2       .       5               TABLE
 7      7     males       DATA      2       .       5               TABLE
 8      8     males$      DATA      2       .       5               TABLE

*            _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| '_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
;


The CONTENTS Procedure

                Directory

Libref             WORK
Engine             V9
Physical Name      d:\wrk\_TD7480_E6420_
Filename           d:\wrk\_TD7480_E6420_
Owner Name         BUILTIN\Administrators
File Size          12KB
File Size (bytes)  12288


              Member   Obs, Entries
 #  Name      Type      or Indexes   Vars  Label     File Size  Last Modified

 1  AGEGE13   DATA          12        5                  128KB  07/22/2020 13:39:09
 2  AGEGE13$  DATA          12        5                  128KB  07/22/2020 13:39:09
 3  AGELE13   DATA          10        5                  128KB  07/22/2020 13:39:09
 4  AGELE13$  DATA          10        5                  128KB  07/22/2020 13:39:09
 5  FEMALES   DATA           9        5                  128KB  07/22/2020 13:39:09
 6  FEMALES$  DATA           9        5                  128KB  07/22/2020 13:39:09
 7  MALES     DATA          10        5                  128KB  07/22/2020 13:39:09
 8  MALES$    DATA          10        5                  128KB  07/22/2020 13:39:09

*
 _ __  _ __ ___   ___ ___  ___ ___
| '_ \| '__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
;

OPTIONS VALIDMEMNAME= EXTEND;
libname xel ("d:/xel/ages.xlsx", "d:/xel/genders.xlsx");
proc copy in=xel out=work;
run;quit;

