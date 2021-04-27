# Email_Valuation_Aggregator
This tool uses Win32com module to connect to my existing outlook session. It then parses my emails based on a date and subject criteria.
After parsing for the specified emails, it downloads the zipfile and decompresses it. The decompression gives me access to the CSV file
inside which I then create a temporary path for using the io.Stringio method. This temp path lets me read the CSV directly into my environment
without an intermediary and I can then perform my dataframe construction and manipulations. It then exports the dataframe results into an excel file
and send email with the attachment and summary tables to my team.
