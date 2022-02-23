<?php

namespace App\Library\Excel\Model;

use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

/**
 * Excel ファイルモデル
 */
class ExcelFile implements WithMultipleSheets
{
    // Excelファイルとして出力できるようになる trait
    use Exportable;

    private array $sheets = [];

    public function addSheets(ExcelSheet $sheet): void
    {
        $this->sheets[] = $sheet;
    }

    /**
     * \Maatwebsite\Excel\Concerns\WithMultipleSheets の指定。
     * 複数のシートから成る Excel ファイルであることを示す
     * @return array
     */
    public function sheets(): array
    {
        return $this->sheets;
    }
}
