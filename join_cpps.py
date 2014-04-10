#! python3
"""
join_cpps: merge costs and descriptions into downloaded table

Usage: join_cpps.py <filename> [-o <output>] [--quiet]  
       join_cpps.py (-h | --help)

Options:
  -h --help    Show this screen.
  <filename>   Download results in filename's list of packages.
  -o <output>  Save as a new file instead of overwriting.
  --quiet      Print less text
"""
import csv
import logging

from docopt import docopt 
import pandas as pd
from find_latest_file import find_latest_file
from config import CPP_DIR # @UnresolvedImport
from config import LIST_DIR # @UnresolvedImport
from config import EFFTYPES # @UnresolvedImport

def get_lasersplit(mc):
    if mc[4:7] == "NOR":
        return "No RD"
    elif EFFTYPES[mc[0]] == "NIns":
        return mc[7:10]
    else:
        return mc[7] + "_" + mc[9]

def get_listcode(mc):
    return "" if EFFTYPES[mc[0]] == "Rnwl" else mc[4:7]

def join_cpps(df):
    
    ## Rename
    df = df.rename(columns={'CLNT': 'Client',
                            'Mail Code': 'Mailcode',
                            'Descript':'IDMI Descript',
                            'Package':'IDMI Package',
                            'Qty Mail': 'Qty Mailed'})
    
    ## And add some columns:
    df['EffType']    = df['Mailcode'].apply(lambda s: EFFTYPES[s[0]])
    df['Eff']        = df['Mailcode'].apply(lambda x: x[:4])
    df["Lasersplit"] = [get_lasersplit(mc) for mc in df['Mailcode']]
    df["Listcode"]  = [get_listcode(mc) for mc in df['Mailcode']]

    ## Drop columns!    
    df = df[['Client','EffType','Eff','Listcode','Lasersplit','Mailcode',
             'IDMI Package','IDMI Descript','Mail Date','FF Date',
             'Qty Mailed','Total Donors','ND Count','Total Revenue']]

    ## Load List Costs
    LCPP_XLSX = find_latest_file('List CPPs', CPP_DIR)
    logging.info('List costs: ' + LCPP_XLSX) 
    lcpps = pd.read_excel(CPP_DIR + LCPP_XLSX, "List & Media CPPs")
    lcpps = lcpps[['Mail Code','Media CPP_ ','Est?']]
    lcpps = lcpps.rename(columns={'Mail Code':'Mailcode',
                                  'Media CPP_ ':'List CPP', 
                                  'Est?':'LCPP Est?'})
    
    logging.info('Merging in list costs...')
    df = pd.merge(df, lcpps, 'left', on='Mailcode')
    
    PCPP_XLSX = find_latest_file('Prod CPPs', CPP_DIR)
    logging.info('Prod costs: ' + PCPP_XLSX) 
    pcpps = pd.read_excel(CPP_DIR + PCPP_XLSX, "Prod CPPs")
    pcpps = pcpps[['Eff', '8&10', 'Total Prd CPP', 'Estimated?']]
    pcpps = pcpps.rename(columns={'8&10':'Lasersplit',
                                  'Total Prd CPP':'Prod CPP', 
                                  'Estimated?':'PCPP Est?'})
    
    ## Merge in the Production Costs
    logging.info('Merging in production costs...')
    df = pd.merge(df, pcpps, 'left', on=['Eff','Lasersplit'])
    
    ## Some nice calulated columns
    df['List Cost'] = df['Qty Mailed'] * df['List CPP']
    df['Prod Cost'] = df['Qty Mailed'] * (df['Prod CPP'])
    
    ## Now we take a bunch of slightly different tabs from a single 
    ## spreadsheet and merge them together (using 'EffType' to distinguish
    ## the originating tabs)

    LIST_XLSX = find_latest_file('USO Prospect Code Log', LIST_DIR)
    logging.info('List descs: ' + LIST_XLSX) 
    
    lists_DM =                 pd.read_excel(LIST_DIR + LIST_XLSX,
                                             "Direct Mail Codes", 
                                             parse_cols="A,B,D,E,F", 
                                             dtypes={'Code'    :'object',
                                                     'Category':'object',
                                                     'List'    :'object',
                                                     'Segment' :'object'
                                                    })
    
    lists_DM = lists_DM.append(pd.read_excel(LIST_DIR + LIST_XLSX,
                                             "USO NY Direct Mail", 
                                             parse_cols="A,B,D,E,F", 
                                             dtypes={'Code'    :'object'}))
    
    lists_DM.rename(columns={'Code'    :'List Code',
                             'Category':'List Category',
                             'List'    :'List Desc',
                             'Segment' :'List Segment/Vehicle'
                            }, inplace=True)
    lists_DM['EffType'] = 'Prsp'
    
    lists_PI    = pd.read_excel(LIST_DIR + LIST_XLSX,
                                "Alternative Media", 
                                parse_cols="A:D", 
                                dtypes={'Code'    :'object'})
    
    lists_PI.rename(columns={'Code'    :'List Code',
                             'Category':'List Category',
                             'Program' :'List Desc',
                             'Details' :'List Segment/Vehicle'
                            }, inplace=True)
    lists_PI['EffType'] = 'PIns'
    
    lists_NI    = pd.read_excel(LIST_DIR + LIST_XLSX,
                                "Newspaper",         
                                parse_cols="A:C", 
                                dtypes={'Code'    :'object'})
    
    lists_NI.rename(columns={'Code'    :'List Code',
                             'Category':'List Category',
                             'Program' :'List Desc'
                            }, inplace=True)
    lists_NI['EffType'] = 'NIns'
    
    lists = pd.concat([lists_DM, lists_PI, lists_NI])
    
    ## Some list codes are inevitably numbers instead of strings, so:
    lists['List Code'] = lists['List Code'].apply(str)
    lists = lists.rename(columns={'List Code': 'Listcode'})
    
    logging.info('Merging in list descriptions...')
    df = pd.merge(df, lists, 'left', on=['EffType','Listcode'])

    return df

def main(args):
    df = pd.read_csv(args['<filename>'])
    df = join_cpps(df)
    outputfilename = args['-o'] if args['-o'] else args['<filename>']
    logging.info('Saving processed results...')
    df.to_csv(outputfilename,
              index=None,
              quoting=csv.QUOTE_NONNUMERIC, 
              date_format='%Y-%m-%d')
    
if __name__ == "__main__":
    args = docopt(__doc__)
    print(args)
    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')

    main(args)