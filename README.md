# Slam PhpSpreadsheet helper to create organized data table

[![Latest Stable Version](https://img.shields.io/packagist/v/slam/openspout-helper.svg)](https://packagist.org/packages/slam/openspout-helper)
[![Downloads](https://img.shields.io/packagist/dt/slam/openspout-helper.svg)](https://packagist.org/packages/slam/openspout-helper)
[![Integrate](https://github.com/Slamdunk/openspout-helper/workflows/Integrate/badge.svg?branch=master)](https://github.com/Slamdunk/openspout-helper/actions)
[![Code Coverage](https://codecov.io/gh/Slamdunk/openspout-helper/coverage.svg?branch=master)](https://codecov.io/gh/Slamdunk/openspout-helper?branch=master)
[![Infection MSI](https://badge.stryker-mutator.io/github.com/Slamdunk/openspout-helper/master)](https://dashboard.stryker-mutator.io/reports/github.com/Slamdunk/openspout-helper/master)

## Installation

`composer require slam/openspout-helper`

## Usage

```php
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

$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();

$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle('My Users');
$table = new ExcelHelper\Table($activeSheet, 1, 1, 'My Heading', $users);
$table->setColumnCollection($columnCollection);

(new ExcelHelper\TableWriter())->writeTable($table);
(new PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet))->save(__DIR__.'/test.xlsx');
```

Result:

![Example](https://raw.githubusercontent.com/Slamdunk/openspout-helper/master/example.png)
