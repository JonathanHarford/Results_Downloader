import pandas as pd

def report_to_csv(df, filename):
    """Convert dates to strings, and make a CSV with 'Mail Code' as the index."""

    def to_iso8601(dt):
        '''Convert a datetime to ISO8601 string.'''
        return dt.strftime('%Y-%m-%d') if (dt > pd.datetime(1900,1,1)) else ''
    
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    # CSV is much faster than XLSX:
    df.to_csv(filename, encoding='utf-8', index_label='Mail Code')