# excel-fill

Tool to automate filling Excel workbook based on predefined rules

## Table of Contents

- [Develop](#develop)
- [Run](#run)
- [Config](#config)

## Develop

### Pre-requisites

- Go 1.25+

### Build

```shell
go build .
```

## Run

```shell
// execute
excel-fill --config samples/config.yaml --in samples/book0.xlsx

// help page
excel-fill --help
```

## Config

Sample `config.yaml` file format.

- `filters` are `OR`-ed
- `actions` are `AND`-ed

```yaml
operations:
  - name: rule-0
    filters:
      - column: A
        value: "filter_val_0"
      - column: B
        value: "^filter_regex_0$"
    actions:
      - column: C
        value: "fill_val_0"
      - column: D
        value: "fill_val_1"
```
