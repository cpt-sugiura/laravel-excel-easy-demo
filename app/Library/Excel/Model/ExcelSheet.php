<?php
// app/Library/Excel/Model/ExcelSheet.php
namespace App\Library\Excel\Model;


use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;

/**
 * Excel シートモデル
 */
class ExcelSheet implements WithTitle, FromArray, WithEvents // WithEvents を追加。これで Laravel-Excel のイベント機能が起動します
{
    /**
     * @param  string  $sheetName  シート名
     */
    public function __construct(
        protected readonly string $sheetName
    ) {
    }

    /**
     * データ本体を格納する。
     * 外から好き勝手操作することが多いのでいっそ public 化。
     * @var array
     */
    public array $rows = [];

    /**
     * シート名。\Maatwebsite\Excel\Concerns\WithTitle の指定
     * @return string
     */
    public function title(): string
    {
        return $this->sheetName;
    }

    /**
     * データ本体。\Maatwebsite\Excel\Concerns\FromArray の指定
     * @return array
     */
    public function array(): array
    {
        return $this->rows;
    }


    /**
     * シートについての処理の最後に発火するイベントのハンドラーを格納するプロパティ
     * @see https://docs.laravel-excel.com/3.1/architecture/#processing-the-sheets-2
     * @var \Closure|null
     */
    public \Closure|null $onAfterSheet = null;

    /**
     * Maatwebsite\Excel の各処理で発火するイベントを登録するメソッド
     *
     * \Maatwebsite\Excel\Concerns\WithEvents のインターフェースの実装
     *
     * @see https://docs.laravel-excel.com/3.1/exports/extending.html#customize
     * @return Closure[]
     */
    public function registerEvents(): array
    {
        // Laravel-Excel に発火させるイベント群
        $events = [];
        if($this->onAfterSheet instanceof \Closure){
            // このクラスに onAfterSheet メソッドがあると IDE が勘違いするコードになるので変数に入れる
            $handleAfterSheet = $this->onAfterSheet;
            // シートについての処理の最後に発火するイベント
            // @see https://docs.laravel-excel.com/3.1/architecture/#processing-the-sheets-2
            $events[AfterSheet::class] = fn(AfterSheet $event) => $handleAfterSheet($event, $this);
        }

        return $events;
    }
}
