<?php

namespace App\Console\Commands;


use App\Library\Excel\Event\AfterSheetHandler;
use App\Library\Excel\Model\ExcelFile;
use App\Library\Excel\Model\ExcelSheet;
use Illuminate\Console\Command;
use Maatwebsite\Excel\Events\AfterSheet;

// php artisan excel-demo で実行
class ExcelDemo extends Command
{
    protected $name = 'excel-demo';

    public function handle(): int
    {
        // ファイルを用意
        $excel = new ExcelFile();
        // シートを用意。ここは Excel を使う目的に応じて色々良く変わります
        $sheet       = new ExcelSheet('シートその1');
        $sheet->rows = [
            ['1行1列目', '浜松太郎'],
            ["複数行を持つセル内の1行目\n複数行の2行目\n3行目"]
        ];

        // シートモデルの中で定義したシートについての AfterSheet イベント（処理の最後に発火するイベント）のハンドラーのプロパティに処理内容を代入
        $sheet->onAfterSheet = static function(AfterSheet $event, ExcelSheet $excelSheet) {
            // 処理内容が複雑なので AfterSheet イベント用のクラスを用意して行いたい処理を呼び出す。
            (new AfterSheetHandler($excelSheet, $event))
                ->wrapText()
                ->adjustColWidth(fontDirPath: 'C:\\Windows\\Fonts\\');

            // Excel を開いた時のカーソルを左上にあわせる
            // 枠線など色々操作するとカーソルが飛ぶ
            $event->sheet->getDelegate()->setSelectedCells("A1:A1");
        };

        // 用意したシートをファイル内に追加
        $excel->addSheets($sheet);
        // ファイルをローカルに保存。 store の代わりに download などもできます。
        // この実体化する時に Laravel-Excel が色々やってくれ、↑の AfterEvent が走ったりもします。
        $excel->store('excelDemo_' . now()->format('YmdHis') . '.xlsx');

        return 0;
    }
}
