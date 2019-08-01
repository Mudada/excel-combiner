# excel-combiner
Script to combine multiple excel worksheet

# Usage
``node dist/main.js -f <path/to/conf/file>``

# Configuration file rules

The configuration file is a JSON file which have those options:

| option          | mandatory | content_type                              | description                                                                                                                                                                                                                                                                                                                                 | content_exemple                                                                                 | content_exemple_explanation                                                                                                                      |
|-----------------|-----------|-------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| files           | *YES*     | [String]                                  | Describe here all the files you want to merge                                                                                                                                                                                                                                                                                               | ["/path/to/file1", "/path/to/file2"]                                                            | Will merge file1 and file2 together                                                                                                              |
| output          | *YES*     | String                                    | Describe here were you want the output to be generated                                                                                                                                                                                                                                                                                      | "/path/to/file_output"                                                                          | Will produce the output as `file_output` in `/path/to` folder                                                                                    |
| start           | NO        | Int                                       | Start of the interval which will be ignored (value included) (useful if you have header for exemple)                                                                                                                                                                                                                                                         | 0                                                                                               | Will not merge the interval starting at row 0 (0 included)                                                                                                    |
| end             | NO        | Int                                       | End of the interval which will be ignored (value excluded) (useful if you have header for exemple)                                                                                                                                                                                                                                                           | 5                                                                                               | Will not merge the interval ending at row 5 (5 excluded)                                                                                                      |
| singleWorksheet | NO        | String | {name: String, except: [String]} | This one is a bit more complex.  If specified as a String it will merge every worksheet as an only worksheet named after the String value. If specified as an object, name will be the worksheet's name and except is a list of worksheet you want to merge "normally" (using there names and only with worksheet that have the same name)  | "Super giga worksheet" | {name: "Super giga worksheet",  except: ["Worksheet1", "Worksheet2"] } | Will merge every worksheet as `Super giga worksheet` | Will merge every worksheet except `Worksheet1` and `Worksheet2` as `Super giga worksheet` |


## Working conf-file exemple 
  ```json
  {
    "files": [
        "/Users/mudada/Code/Script/excel-combiner/excel/Tableau Carnot TSN-EP-v3.xlsx",
        "/Users/mudada/Code/Script/excel-combiner/excel/Tableau Carnot TSN-Eurecom-v3.xlsx"
    ],
    "singleWorksheet": {
        "name": "Super Worksheet"
        , "except": [
            "00-Definitions",
            "11-Référentiel"
        ]
    },
    "output": "/Users/mudada/Code/Script/excel-combiner/excel/output/output.xlsx",
    "start": 0,
    "end": 5
}
```
