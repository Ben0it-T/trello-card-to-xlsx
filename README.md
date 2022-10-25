# Trello Card To Xlsx

Converts Trello Card JSON export to XLSX file.

## Requirements

- Python 3.6+
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

```
python3 trelloCardToXlsx.py card.json 
```

## Notes

With a Trello 'free plan' can only export a card to JSON.
For my personnal use, I wrote this python script to convert a JSON card to an XLSX file (that I can archive to document my projects for example)

## Example

<img width="373" alt="Trello-Convert-Card-JSON-to-XLS" src="https://user-images.githubusercontent.com/37017213/197360018-465ee4ba-9e85-46af-9d87-ebaa178a1945.png">

