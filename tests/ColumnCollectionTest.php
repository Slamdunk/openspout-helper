<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\Tests;

use PHPUnit\Framework\TestCase;
use Slam\OpenspoutHelper\CellStyle\Text;
use Slam\OpenspoutHelper\Column;
use Slam\OpenspoutHelper\ColumnCollection;

final class ColumnCollectionTest extends TestCase
{
    private Column $column;
    private ColumnCollection $collection;

    protected function setUp(): void
    {
        $this->column     = new Column('foo', 'Foo', 10, new Text());
        $this->collection = new ColumnCollection(...[$this->column]);
    }

    public function testBaseFunctionalities(): void
    {
        self::assertArrayHasKey('foo', $this->collection->getArrayCopy());
        self::assertSame($this->column, $this->collection['foo']);
    }
}
