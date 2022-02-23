<?php

namespace App\Library\Excel\Event;

use App\Library\Excel\Model\ExcelSheet;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Helper\Dimension;
use PhpOffice\PhpSpreadsheet\Shared\Font as FontShared;
use PhpOffice\PhpSpreadsheet\Style\Font as FontStyle;

/**
 * Laravel-Excel の AfterSheet イベントを操作する色々。
 */
class AfterSheetHandler
{
    /**
     * @param  ExcelSheet  $excelSheet  アプリケーション内で用意したシートモデル
     * @param  AfterSheet  $event       Laravel-Excel で用意済みのシートの諸々完了後に発火するイベント
     */
    public function __construct(
        protected ExcelSheet $excelSheet,
        protected AfterSheet $event
    ) {
    }

    /**
     * 渡された範囲のセルの"折り返して全体を表示する"をONにする
     * @param  string|null  $range  ex. A1:Z100
     * @return $this
     */
    public function wrapText(string $range = null): static
    {
        // 折り返し範囲を決定する。外から与えられなければセルが埋まっている範囲を対象にする
        $range ??= $this->getRangeOfIncludeAllNotEmptyCell();
        if($range === null) {
            //  埋まっているセルが存在しないのであれば折り返しを一切作らずに終わり
            return $this;
        }

        // 対象のセルらを折り返しありにする。
        $this->event->sheet->getDelegate() // getDelegate で \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet インスタンスを呼び出す
        ->getStyle($range) // ここから PhpSpreadsheet のメソッド
        ->getAlignment()
            ->setWrapText(true);

        return $this;
    }

    /**
     * 値が入っているセル全てを含む範囲を示す文字列を返す。
     * 値が入っているセルが全くない場合は null を返す。
     * @return string|null
     */
    protected function getRangeOfIncludeAllNotEmptyCell(): ?string
    {
        // 値が入っているセルら上下左右の端のインデックスを得る
        $conner = $this->getConnerIndexes();
        if(in_array(null, $conner, true)) {
            //  埋まっているセルが存在しないのであれば折り返しを一切作らずに終わり
            return null;
        }
        // 得られた端の列インデックスを Excel の A,B,... に変換する
        $colIndexMin = Coordinate::stringFromColumnIndex($conner['colIndexMin'] + 1);
        $colIndexMax = Coordinate::stringFromColumnIndex($conner['colIndexMax'] + 1);

        // 得られた端のインデックスから範囲定義文字列を構築
        return $colIndexMin . ($conner['rowIndexMin'] + 1)
            . ':' . $colIndexMax . ($conner['rowIndexMax'] + 1);
    }

    /**
     * 値のあるセルの最小行、最大行、最小列、最大列を得る
     * @return array{rowIndexMin: int|null, rowIndexMax: int|null , colIndexMin: int|null , colIndexMax: int|null }
     */
    protected function getConnerIndexes(): array
    {
        // セルの値を二次元配列化
        $cells = $this->event->sheet->getDelegate()->toArray();
        // undefined error 防止に null で初期化
        $rowIndexMin = $rowIndexMax = $colIndexMin = $colIndexMax = null;
        // 行ループ
        foreach($cells as $rowIndex => $row) {
            // 行が空でない場合のみ処理を行う
            if(!empty(array_filter($row, static fn($cell) => $cell !== null))) {
                $rowIndexMin ??= $rowIndex; // 最初の一度だけ代入
                $rowIndexMax = $rowIndex; // 空でない行が現れ次第、更新
            }

            // 列ループ
            // ループ回数削減のため、現在の最大値より右だけを探索する
            $colIndexStart = isset($colIndexMax) ? ($colIndexMax + 1) : 0;
            $colIndexEnd   = count($row) - 1;
            for($colIndex = $colIndexStart; $colIndex <= $colIndexEnd; $colIndex++) {
                // セルが空で内場合のみ処理を行う
                if(isset($row[$colIndex])) {
                    $colIndexMin ??= $colIndex; // 最初の一度だけ代入
                    $colIndexMax = $colIndex; // 現在の最大値より右に空でないセルが見つかり次第、更新
                }
            }
        }

        return compact('rowIndexMin', 'rowIndexMax', 'colIndexMin', 'colIndexMax');
    }

    /**
     * 一時的にフォントファイル格納ディレクトリのパス設定を上書きして処理を行う。
     * フォントディレクトリパス設定が PhpSpreadSheet 内全体で共通設定なため欲しくなる処理。
     * @param  string    $fontDirPath  使うフォントファイル格納ディレクトリのパス
     * @param  callable  $callback     処理内容
     * @return AfterSheetHandler
     */
    public function runCallbackWithTemporaryFontPath(string $fontDirPath, callable $callback): static
    {
        // 一時的にフォント参照場所の設定を変えるので一時退避
        $fontPathWhenCalled = FontShared::getTrueTypeFontPath();
        try {
            // PhpSpreadSheet はフォントディレクトリパスの末尾スラッシュが抜けていると期待しない挙動を示すので
            // 末尾スラッシュがあるべきならばここで付与
            if(!preg_match('#/$#', $fontDirPath)) {
                $fontDirPath .= '/';
            }
            // フォントファイルらが格納されているディレクトリについてのパスを設定
            FontShared::setTrueTypeFontPath($fontDirPath);

            $callback();
        } finally {
            // フォント参照場所を復元
            FontShared::setTrueTypeFontPath($fontPathWhenCalled);
        }

        return $this;
    }

    /**
     * 渡された範囲のセルを元に列の幅を文字が収まるぎりぎりの幅に調整する
     * @param  string|null  $range        ex. A1:Z100
     * @param  string       $fontDirPath  幅を測るために使うフォントファイルらが保存されているのディレクトリのパス。Windows上で使うなら C:\\Windows\\Fonts\\ がおすすめ
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function adjustColWidth(string $range = null, string $fontDirPath = __DIR__ . '/../fonts/'): AfterSheetHandler
    {
        // 一時的にフォントファイル設定を書き変える。グローバルな設定のためこの形
        return $this->runCallbackWithTemporaryFontPath($fontDirPath, function() use ($range) {
            // 幅調整範囲を決定する。外から与えられなければセルが埋まっている範囲を対象にする
            $range ??= $this->getRangeOfIncludeAllNotEmptyCell();
            if($range === null) {
                //  埋まっているセルが存在しないのであれば幅調整を一切作らずに終わり
                return;
            }
            // 対象範囲文字列を分解して行、列をどこからどこまで参照すればよいか変数に格納
            preg_match('/([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)/', $range, $matches);
            $colStartIndex = strtoupper($matches[1]);
            $rowStartIndex = $matches[2];
            $colEndIndex   = strtoupper($matches[3]);
            $rowEndIndex   = $matches[4];

            // 対象範囲のセルを総ざらいして調整した幅を設定する
            foreach(range($colStartIndex, $colEndIndex) as $colIndex) {
                /** @var float[] $cellFitWidthList セルに合った幅をまとめた配列。最後はこの中で最大のものを列幅とする */
                $cellFitWidthList = [];
                foreach(range($rowStartIndex, $rowEndIndex) as $rowIndex) {
                    // このループで参照しているセルがいい感じに収まる幅を格納
                    $cellFitWidthList[] = $this->getCellFitWidth($colIndex . $rowIndex);
                }
                // 集めた幅の中から最大のものを列幅にセット
                $this->event->getDelegate()->getColumnDimension($colIndex)->setWidth(max([
                    PHP_INT_MIN,// 番人。空配列を渡すことによる max 関数の引数エラー防止。マイナスを PhpSpreadSheet に渡すと特に処理せず自然に進めてくれる
                    ...$cellFitWidthList
                ]));
            }
        });
    }

    /**
     * あるセルがピッタリ収まる幅を取得する
     * @param  string  $coordinate  セルの番地。例: A1
     * @return float|null
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function getCellFitWidth(string $coordinate): ?float
    {
        // シートについて色々細かく操作できる \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet インスタンスを
        // Laravel 用にラッピングされた \Maatwebsite\Excel\Sheet の中から呼び出す
        $officeSheet = $this->event->sheet->getDelegate();

        /** @var string|null $cellValue セルの中に入っている値。文字列か空か */
        $cellValue = $officeSheet->getCell($coordinate)->getValue();
        if($cellValue === null) {
            // セルの中身がないのであればピッタリの幅は定義できないとして null を返す
            return null;
        }

        /** @var FontStyle $font 対象セルで使っているフォント設定 */
        $font = $officeSheet->getStyle($coordinate)->getFont();
        /** @var string $fontFilePath 対象セルで使っているフォントファイルのパス。PhpSpreadsheetの用意した静的メソッドから取得 */
        $fontFilePath = FontShared::getTrueTypeFontFileFromFont($font);

        /** @var float[] $widthList 複数行がセル内にある場合を考慮して各行の幅を考える */
        $widthList = [];
        foreach(explode("\n", $cellValue) as $line) {// セルの中のテキストを行に分解
            // @see https://www.php.net/manual/ja/function.imagettfbbox.php
            // imagettfbbox 関数でフォントから文字が占有する大きさを取得

            // パディングの長さ。PhpSpreadsheet のコードをより平の PHP に近づけた感じ
            // なぜ n なのか、なぜ 1.07 なのか、なぜ ceil なのかは不明
            // @see vendor/phpoffice/phpspreadsheet/src/PhpSpreadsheet/Shared/Font.php:250
            $textBox = imagettfbbox($font->getSize(), 0, $fontFilePath, 'n');
            $padding = ceil(($textBox[2] - $textBox[0]) * 1.07);

            // 本文の長さ
            $textBox = imagettfbbox($font->getSize(), 0, $fontFilePath, $line);
            $body    = $textBox[2] - $textBox[0];
            // PhpSpreadsheet 内に入っている定数でフォント界のポイント単位と Excel 界のミリメートル単位を変換する
            $widthList[] = ($body + $padding) / Dimension::ABSOLUTE_UNITS[Dimension::UOM_MILLIMETERS];
        }
        return max([
            PHP_INT_MIN,// 番兵。空の配列を max 関数に渡すとエラーになるのでそれの防止。マイナスを PhpSpreadSheet に渡すと特に処理せず自然に進めてくれる
            ...$widthList
        ]);
    }
}
