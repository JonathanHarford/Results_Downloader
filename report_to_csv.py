#! python3
import pandas as pd

def report_to_csv(df, filename):
    """Convert dates to strings, and make a CSV."""

    def to_iso8601(dt):
        '''Convert a datetime to ISO8601 string.'''
        return dt.strftime('%Y-%m-%d') if (dt > pd.datetime(1900,1,1)) else ''
    
    # Dates are ugly unless we do this:
    for col in df.columns:
        if df[col].dtype.str == '<M8[ns]':
            df[col]  = df[col].apply(to_iso8601)
        
    # CSV is much faster than XLSX:
    df.to_csv(filename + '.csv', encoding='utf-8', index_label=df.index.names)