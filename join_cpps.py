#! python3
"""
join_cpps: merge costs and descriptions into downloaded table

Usage: join_cpps.py <filename> [-o OUTPUT] [--quiet]  
       join_cpps.py (-h | --help)

Options:
  -h --help    Show this screen.
  <filename>   Download results in filename's list of packages.
  -o=<output>  Save as a new file instead of overwriting.
  --quiet      Print less text
"""

import logging

from docopt import docopt 
import pandas as pd
from find_latest_file import find_latest_file
from report_to_csv import report_to_csv
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
    lcpps = pd.read_excel(CPP_DIR + LCPP_XLSX, "List & Media CPPs", parse_cols="A,K,L") # G is List Desc
    
    logging.info('Merging in list costs...') 
    df = df.set_index(['Mailcode']).join(lcpps.set_index(['Mail Code']))
    df = df.reset_index()
    
    ## Rename some of the imported columns
    df = df.rename(columns={'Media CPP_ ':'List CPP', 'Est?':'LCPP Est?'})

    PCPP_XLSX = find_latest_file('Prod CPPs', CPP_DIR)
    logging.info('Prod costs: ' + PCPP_XLSX) 
    pcpps = pd.read_excel(CPP_DIR + PCPP_XLSX, "Prod CPPs", parse_cols='C,E,M,P')
    
    ## Merge in the Production Costs
    logging.info('Merging in production costs...')
    df = df.set_index(['Eff','Lasersplit']).join(pcpps.set_index(['Eff','8&10']))
    df = df.reset_index()
    
    ## Rename some of the imported columns
    df = df.rename(columns={'Total Prd CPP':'Prod CPP', 'Estimated?':'PCPP Est'})

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
    
    logging.info('Merging in list descriptions...')
    df = df.set_index(['EffType','Listcode']).join(lists.set_index(['EffType','List Code']))
    # df = df.reset_index()

    return df

def main(args):
    df = pd.read_csv(args['<filename>'])
    df = join_cpps(df)
    outputfilename = args['-o'] if args['-o'] else args['<filename>']
    logging.info('Saving processed results...')
    report_to_csv(df.set_index("Mailcode"), outputfilename)

    
if __name__ == "__main__":
    args = docopt(__doc__)
    print(args)
    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')

    main(args)