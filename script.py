import requests
import wget
import os
from glob import glob
import pandas as pd
from tqdm import tqdm

df = pd.read_excel("books.xlsx")
number_of_books = len(df)
categories = df['English Package Name'].unique()
print(
    f'There are {number_of_books} books available from {len(categories)} categories.')

books_downloaded = list(
    map(os.path.basename, glob('**/*.pdf', recursive=True)))
print(f'{len(books_downloaded)} of {number_of_books} already downloaded.')

# Create subdirectories
for c in categories:
    if not os.path.exists(c):
        os.mkdir(c)

# Obtain all books in the current directory
for index, row in tqdm(df.iterrows()):
    file_name = f"{row.loc['Book Title']}_{row.loc['Edition']}".replace(
        '/', '-').replace(':', '-')
    if f'{file_name}.pdf' not in books_downloaded:
        print(f'\nDownloading {file_name}')
        url = f"{row.loc['OpenURL']}"
        r = requests.get(url)
        download_url = f"{r.url.replace('book', 'content/pdf')}.pdf"
        wget.download(download_url, f"{file_name}.pdf")

    # Move book into subdirectory
    if os.path.exists(f'{file_name}.pdf'):
        os.rename(f'{file_name}.pdf',
                  f"{row.loc['English Package Name']}/{file_name}.pdf")
else:
    print('Finished downloading.')
