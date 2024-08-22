import pandas as pd
import os

def get_filepath(directory: str) -> str:
    files = [os.path.join(directory, file) for file in os.listdir(directory) if os.path.isfile(os.path.join(directory, file))]
    
    return max(files, key=os.path.getctime, default=None)

def load_funds_available(filepath: str) -> pd.DataFrame:
    return pd.read_excel(filepath)

def determine_entry(user_input: str) -> str:
    if '.' in user_input:
        return 'string'
    else:
        return 'fund'

def retrieve_funds_available(glString: str,dataframe: pd.DataFrame) -> tuple[str,str,str]:
    parent = dataframe.loc[dataframe['GL Account'] == glString]['parentAcct'].values[0]

    acct_type = dataframe.loc[dataframe['GL Account'] == glString]['acctType'].values[0]
    funds = dataframe.loc[dataframe['GL Account'] == parent]['Funds Available'].values[0]
    

    return acct_type,parent,funds

def lookup_interface(dataframe: pd.DataFrame) -> str:
    while True:
        inp = input('Enter Fund or Full GL String: ')
        match determine_entry(inp):
            case 'string':
                gl = inp
            case 'fund':
                gl = [input(f'Enter {seg}: ') for seg in ['Office','Program','Account']]
                gl.insert(0,inp)
                gl = str(gl).lstrip('[').rstrip(']').replace(', ','.').replace("'","")
            case _:
                return 'Unknown Entry Type'
        
        gl += '.00000.00000'

        if len(gl) != len('fffff.oooo.ppppp.aaaaaa.00000.00000'):
            print(f'\nInvalidGlString: {gl}\n')
            continue
        else:
            break

    acct_type,parent,funds = retrieve_funds_available(gl,dataframe)

    text = f'\nEntered Account: {gl} ({acct_type})\nParent Account:  {parent}\n Available: {float(funds):,}\n'
    
    return text

def main() -> None:
    rep_dir = r'C:\Users\nathansmalley\OneDrive - Cook County Government\1 - Reports\FundsAvailable'
    most_recent = get_filepath(rep_dir)
    df = load_funds_available(most_recent)

    print(lookup_interface(df))

if __name__ == '__main__':
    main()