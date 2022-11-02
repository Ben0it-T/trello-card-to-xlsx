# Trello Card To Xlsx

Converts Trello Card JSON export to XLSX file.

With a Trello 'free plan' we can only export a card to JSON.
For my personnal use, I wrote this first python script to convert a JSON card to an XLSX file.

## Requirements

- Python 3
- XlsxWriter (https://pypi.org/project/XlsxWriter/)

## config.ini

`config.ini` is a very basic configuration file. Il contains three sections :
- `[Dates]`: time zone and date format
- `[Labels]`: custom titles
- `[TrelloLists]`: map list `id` and list `name`

```
[TrelloLists]
<list id> = <list name>
```

## Usage

Export your card to json

Run
```
python3 trelloCardToXlsx.py card.json 
```

## Example

<img width="373" alt="Trello-Convert-Card-JSON-to-XLS" src="https://user-images.githubusercontent.com/37017213/197360018-465ee4ba-9e85-46af-9d87-ebaa178a1945.png">

[Trello-Convert-Card-JSON-to-XLS.xlsx](https://github.com/Ben0it-T/trello-card-to-xlsx/files/9857147/Trello-Convert-Card-JSON-to-XLS.xlsx)
