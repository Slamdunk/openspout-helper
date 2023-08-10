# Slam PhpSpreadsheet helper to create organized data table

[![Latest Stable Version](https://img.shields.io/packagist/v/slam/openspout-helper.svg)](https://packagist.org/packages/slam/openspout-helper)
[![Downloads](https://img.shields.io/packagist/dt/slam/openspout-helper.svg)](https://packagist.org/packages/slam/openspout-helper)
[![CI](https://github.com/Slamdunk/openspout-helper/actions/workflows/ci.yaml/badge.svg)](https://github.com/Slamdunk/openspout-helper/actions/workflows/ci.yaml)
[![Infection MSI](https://img.shields.io/endpoint?style=flat&url=https%3A%2F%2Fbadge-api.stryker-mutator.io%2Fgithub.com%2FSlamdunk%2Fopenspout-helper%2Fmain)](https://dashboard.stryker-mutator.io/reports/github.com/Slamdunk/openspout-helper/main)


## Installation

`composer require slam/openspout-helper`

## Usage

```php
use OpenSpout\Writer\Common\Creator\WriterEntityFactory;
use Slam\OpenspoutHelper as ExcelHelper;

require __DIR__ . '/vendor/autoload.php';

// Being an `iterable`, the data can be any dinamically generated content
// for example a PDOStatement set on unbuffered query
$users = [
    [
        'column_1' => 'John',
        'column_2' => '123.45',
        'column_3' => '2017-05-08',
    ],
    [
        'column_1' => 'Mary',
        'column_2' => '4321.09',
        'column_3' => '2018-05-08',
    ],
];

$columnCollection = new ExcelHelper\ColumnCollection(...[
    new ExcelHelper\Column('column_1',  'User',     10,     new ExcelHelper\CellStyle\Text()),
    new ExcelHelper\Column('column_2',  'Amount',   15,     new ExcelHelper\CellStyle\Amount()),
    new ExcelHelper\Column('column_3',  'Date',     15,     new ExcelHelper\CellStyle\Date()),
]);

$XLSXWriter  = WriterEntityFactory::createXLSXWriter();
$XLSXWriter->openToFile(__DIR__.'/test.xlsx');

$activeSheet = $XLSXWriter->getCurrentSheet();
$activeSheet->setName('My Users');
$table = new ExcelHelper\Table($activeSheet, 'My Heading', $users);
$table->setColumnCollection($columnCollection);

(new ExcelHelper\TableWriter())->writeTable($XLSXWriter, $table);
$XLSXWriter->close();
```

Result:

![Example](https://raw.githubusercontent.com/Slamdunk/openspout-helper/master/example.png)
