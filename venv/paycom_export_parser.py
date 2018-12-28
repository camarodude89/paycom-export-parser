from pandas import DataFrame, read_excel
import pandas as pd
import re


def mc_and_mac_fix(s):
    pos = re.search('Mc|Mac', s).end()
    char = s[pos]
    string_list = list(s)
    string_list[pos] = char.swapcase()
    return ''.join(string_list)


def ordinal_suffix_fix(s):
    string_list = re.split('\s', s)

    if re.match("^(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$", string_list[-1], flags=re.IGNORECASE):
        string_list.append(string_list.pop().upper())
        return ' '.join(string_list)

    return ' '.join(string_list)


def main():
    file = r'20181207070755_Active Team members_.xls'
    df = pd.read_excel(file)
    df['Employee_Name'] = df['Employee_Name'].str.title()

    # Correctly capitalizes all names with 'Mc' and 'Mac' in them such as 'McCoy' or 'MacDonald'
    mc_and_mac_name_col = df[df['Employee_Name'].str.contains('Mc|Mac')]['Employee_Name']
    mc_and_mac_list = mc_and_mac_name_col.tolist()
    corrected_mc_and_mac_list = list(map(mc_and_mac_fix, mc_and_mac_list))
    df['Employee_Name'].replace(to_replace=mc_and_mac_list, value=corrected_mc_and_mac_list, inplace=True)

    df[['Last_Name', 'First_Middle_Names']] = df['Employee_Name'].str.split(',\s', expand=True)

    # Increases the columns that are visible when printing
    pd.set_option('display.max_columns', 500)

    df.drop('Employee_Name', axis=1, inplace=True)
    df = pd.concat([df[['Employee_Code']], df['First_Middle_Names'].str.split('\s', expand=True, n=1),
                    df[['Last_Name']]], axis=1)
    df.rename(columns={0: 'First_Name', 1: 'Middle_Name'}, inplace=True)
    df['User_ID'] = df['Employee_Code'] + 10000
    df.drop('Employee_Code', axis=1, inplace=True)

    # Correctly capitalizes all names with ordinal suffixes such as III or VI
    spaced_last_name_col = df[df['Last_Name'].str.contains('\s')]['Last_Name']
    spaced_last_name_list = spaced_last_name_col.tolist()
    corrected_last_name_list = list(map(ordinal_suffix_fix, spaced_last_name_list))
    df['Last_Name'].replace(to_replace=spaced_last_name_list, value=corrected_last_name_list, inplace=True)

    df.to_csv('Paycom_Parsed_Export.csv', index=False)

if __name__ == "__main__":
    main()