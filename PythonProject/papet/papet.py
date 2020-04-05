
#!/usr/bin/python3
import os
import sys
import pandas as pd
from argparse import ArgumentParser


def get_args():
    """
     * Get configuration for program execution
    """
    parser = ArgumentParser()

    # IO
    parser.add_argument(
        '-s', '--sedona', default=None, type=str,
        help='Fisierul cu date pentru Sedona.')

    parser.add_argument(
        '-g', '--saga', default=None, type=str,
        help='Fisierul cu date pentru Saga.')

    parser.add_argument(
        '-f', '--furnizori', default='furnizori.xls', type=str,
        help='Fisierul cu date pentru furnizori.')

    parser.add_argument(
        '-r', '--rezultat', default='aprovizionare.xlsx', type=str,
        help='Fisierul in care se vor scrie rezultate.')

    return parser.parse_args()


def process(args):
    """
     * Find items in Saga and Sedona that need to be ordered.
     * Also note internal code errors and items without suppliers.
    """
    def fix_cod(x):
        try:
            ix = int(x)
            return "%d" % ix
        except:
            return x

    errors = []
    # read furnizori
    fdf = pd.read_excel(args.furnizori, converters={'cod': str,
                                                    'denumire': str,
                                                    'um': str,
                                                    'den_tip': str,
                                                    'furnizor': str})
    fdf['cod'] = fdf['cod'].apply(fix_cod)

    # find duplicate codes in furnizori
    dups = fdf[fdf.duplicated(subset=['cod'], keep='last')]['cod'].tolist()
    if dups:
        errors.append('Coduri interne duplicate in furnizori care au fost agregate: %s. ' % repr(dups))
        fdf = fdf.groupby('cod').agg({'denumire': 'first',
                                      'um': 'first',
                                      'den_tip': 'first',
                                      'furnizor': 'first',
                                      'cantit minima': 'min',
                                      'cantit maxima': 'max'}).reset_index()

    # read saga
    sdf = pd.read_excel(args.saga, converters={'cod': str,
                                               'denumire': str,
                                               'um': str,
                                               'den_tip': str})
    sdf['cod'] = sdf['cod'].apply(fix_cod)


    # find duplicate codes in saga
    dups = sdf[sdf.duplicated(subset=['cod'], keep='last')]['cod'].tolist()
    if dups:
        errors.append('Coduri interne duplicate in saga care au fost agregate: %s. ' % repr(dups))
        sdf = sdf.groupby('cod').agg({'denumire': 'first',
                                      'um': 'first',
                                      'den_tip': 'first',
                                      'stoc': 'sum'}).reset_index()

    # read sedona
    ddf = pd.read_excel(args.sedona, converters={'Departament': str,
                                                 'Produs': str,
                                                 'Cod intern': str,
                                                 'U.M.': str})
    ddf['Cod intern'] = ddf['Cod intern'].apply(fix_cod)
    # rename sedona columns to match saga
    ddf = ddf.rename(columns={'Cod intern':'cod'})
    # find duplicate codes in sedona
    dups = ddf[ddf.duplicated(subset=['cod'], keep='last')]['cod'].tolist()
    if dups:
        errors.append('Coduri interne duplicate in sedona care au fost agregate: %s. ' % repr(dups))
        ddf = ddf.groupby('cod').agg({'Departament': 'first',
                                      'Produs': 'first',
                                      'Cod de bare': 'first',
                                      'PLU': 'first',
                                      'U.M.': 'first',
                                      'Cota TVA': 'max',
                                      'Stoc curent': 'sum',
                                      'Ultimul pret de achizitie fara TVA': 'max',
                                      'Valoare achizitie fara TVA': 'max',
                                      'Adaos': 'max',
                                      'Adaos %': 'max',
                                      'Pret fara TVA': 'max',
                                      'Pret cu TVA': 'max'}).reset_index()

    # join the three sheets based on cod
    jdf = sdf.set_index('cod').join(ddf.set_index('cod')).join(fdf.set_index('cod'), rsuffix='_f')

    # add rows from sedona that don't exist in saga
    dd = ddf[~ddf['cod'].isin(sdf['cod'])].rename(columns={'Produs': 'denumire', 'U.M.': 'um'})
    dd = dd.set_index('cod').join(fdf.set_index('cod'), rsuffix='_f')
    jdf = jdf.append(dd, sort=False)

    # keep only columns of interest
    jdf = jdf[['denumire', 'um', 'den_tip', 'stoc', 'Stoc curent',
               'Ultimul pret de achizitie fara TVA', 'furnizor',
               'cantit minima', 'cantit maxima'
               ]]

    # rename columns and add composite stoc column
    jdf = jdf.rename(columns={'stoc': 'stoc_saga', 'Stoc curent': 'stoc_sedona'})
    jdf['stoc_saga'] = jdf['stoc_saga'].fillna(0.0)
    jdf['stoc_sedona'] = jdf['stoc_sedona'].fillna(0.0)
    jdf.insert(5, 'stoc', jdf['stoc_saga'] + jdf['stoc_sedona'])

    # keep rows that need stoc
    stoc = jdf[jdf['stoc'] < jdf['cantit minima']].copy().sort_values(['furnizor', 'cod'])
    # add aprovizionat column
    stoc['aprovizionat'] = stoc['cantit maxima'].sub(stoc['stoc'], axis=0)

    # find rows without furnizor
    furn = jdf[jdf['furnizor'].isnull()].copy()

    # transform errors into dataframe
    errors = pd.DataFrame(errors, columns=['Eroare'])

    # write output spreadsheet
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(args.rezultat, engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    stoc.to_excel(writer, sheet_name='Aprovizionare')
    furn.to_excel(writer, sheet_name='Fara_furnizori')
    errors.to_excel(writer, sheet_name='Erori', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


if __name__ == '__main__':
    args = get_args()

    if args.sedona is None or not os.path.isfile(args.sedona):
        sys.exit("No am gasit fisierul de date pentru Sedona.")

    if args.saga is None or not os.path.isfile(args.saga):
        sys.exit("No am gasit fisierul de date pentru Saga.")

    if args.furnizori is None or not os.path.isfile(args.furnizori):
        sys.exit("No am gasit fisierul de date pentru Furnizatori.")

    process(args)

# def test(args):
#     return print(args)
