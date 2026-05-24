# RVTools-trimmer
Just keep the (very) important parts of the RVTools report

This is a quick'n'dirty script to remove excess fat from a RVtools report file. 
I did that so my tiny AI system can use and process it.

Just do a 
python3 -m venv venv
source venv/bin/activate
pip -r requirements.txt

In order to run the script, do a 

python clean_rvtools.py <rvtools-report.xlsx>

It works with xlsx created with latest rvtools report utility

