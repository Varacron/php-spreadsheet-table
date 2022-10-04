<?php
// TODO: row class
// TODO: row attribute
// TODO: cell attribute
// TODO: thead, tbody, tfoot
namespace Varacron\SpreadsheetTable;

class SpreadsheetTable
{
    private $intentCount = 2;
    private $beautify = 2;

    private $colLetters = [];
    private $colFrom = 'A';
    private $colTo = null;
    private $colSkips = [];

    private $rowFrom = 1;
    private $rowTo = null;
    private $rowSkips = [];

    private $classCellDefault = [];
    private $classCell = [];

    private $cellValueFormatter = [];

    private $spreadsheet;
    private $mergedCells;
    private $tempRows;
    private $tempCols;

    public function __construct()
    {
        $az = range('A', 'Z');
        $azc = count($az);

        $this->colLetters = $az;
        for ($l1 = 0; $l1 < $azc; $l1++) {
            for ($l2 = 0; $l2 < $azc; $l2++) {
                $this->colLetters[] = $az[$l1] . $az[$l2];
            }
        }
    }

    public function load($file_path)
    {
        if (!file_exists($file_path)) {
            throw new \Illuminate\Contracts\Filesystem\FileNotFoundException('SpreadsheetTable file not found: ' . $file_path);
        }
        $this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
        $this->calculateMergedCells();
        return $this;
    }

    private function calculateMergedCells()
    {
        $this->mergedCells = [];
        $temp = $this->spreadsheet->getActiveSheet()->getMergeCells();
        foreach ($temp as $c) {
            $n = [];
            $tt = explode(':', $c, 2);
            foreach ($tt as $t) {
                preg_match('/([A-Z]+)(\d+)/', $t, $m);
                $n[] = $m[1];
                $n[] = $this->colLetterVal($m[1]);
                $n[] = (int) $m[2];
            }
            // $cells = [];
            // $fc = array_search($n[0], $this->colLetters);
            // $lc = array_search($n[3], $this->colLetters);
            // for ($ri = $n[2]; $ri <= $n[5]; $ri++) {
            //     for ($ci = $fc; $ci <= $lc; $ci++) {
            //         $cells[] = $this->colLetters[$ci] . $ri;
            //     }
            // }
            // $n[] = $cells;
            $n['range'] = $c;
            $this->mergedCells[] = $n;
        }
    }

    public function changeSheetIndex($index)
    {
        $this->spreadsheet->setActiveSheetIndex($index);
        $this->calculateMergedCells();
    }

    public function setColumnRange($from, $to)
    {
        $this->colFrom = $from;
        $this->colTo = $to;
    }

    public function setRowMin($value)
    {
        $this->rowFrom = $value;
    }

    public function setRowMax($value)
    {
        $this->rowTo = $value;
    }

    public function setRowRange($min, $max)
    {
        $this->rowFrom = $min;
        $this->rowTo = $max;
    }

    public function setRange($range)
    {
        $tt = explode(':', $range, 2);
        $n = [];
        foreach ($tt as $t) {
            preg_match('/([A-Z]+)(\d+)/', $t, $m);
            $n[] = $m[1];
            $n[] = (int) $m[2];
        }
        $this->setRowRange($n[1], $n[3]);
        $this->setColumnRange($n[0], $n[2]);
    }

    public function addRowSkip()
    {
        $args = func_get_args();
        foreach ($args as $arg) {
            if (is_array($arg)) {
                foreach ($arg as $arg_) {
                    $this->rowSkips[] = $arg_;
                }
            } else {
                $this->rowSkips[] = $arg;
            }
        }
    }

    public function addColSkip()
    {
        $args = func_get_args();
        foreach ($args as $arg) {
            if (is_array($arg)) {
                foreach ($arg as $arg_) {
                    $this->colSkips[] = $arg_;
                }
            } else {
                $this->colSkips[] = $arg;
            }
        }
    }

    public function setCellDefaultClass()
    {
        $args = func_get_args();
        foreach ($args as $arg) {
            if (is_array($arg)) {
                foreach ($arg as $arg_) {
                    $this->classCellDefault[] = $arg_;
                }
            } else {
                $this->classCellDefault[] = $arg;
            }
        }
    }

    public function setCellClass()
    {
        $args = func_get_args();
        $argsCount = count($args);
        $ranges = $args[0];
        if (!is_array($ranges)) {
            $ranges = [$ranges];
        }
        unset($args[0]);
        $classStart = 1;
        $type = 1;
        if ($argsCount > 1) {
            if (in_array($args[1], [0, 1, -1], true)) {
                $type = $args[1];
                unset($args[1]);
                $classStart++;
            }
        }
        $args = array_values($args);
        foreach ($ranges as $range) {
            $this->classCell[] = [$range, $type, $args];
        }
    }

    public function setBeautify($value)
    {
        $this->beautify = $value;
    }

    public function setDefaultIntent($value)
    {
        $this->intentCount = $value;
    }

    private function intent($plus)
    {
        return str_repeat("\t", $plus + $this->intentCount);
    }

    private function renderRow($rowIndex)
    {
        $return = '';
        if ($this->beautify > 0) {
            $return .= "\n" . $this->intent(2);
        }
        $return .= "<tr>";
        foreach ($this->tempCols as $colLetter) {
            $return .= $this->renderCell($rowIndex, $colLetter);
        }
        if ($this->beautify > 1) {
            $return .= "\n" . $this->intent(2);
        }
        $return .= '</tr>';
        return $return;
    }

    private function renderCell($rowIndex, $colLetter)
    {
        $cell = $this->spreadsheet->getActiveSheet()->getCell($colLetter . $rowIndex);
        $return = '';
        if ($this->beautify > 1) {
            $return .= "\n" . $this->intent(3);
        }
        $return .= '<td';
        $colLetterVal = $this->colLetterVal($colLetter);
        $colSpan = 1;
        $rowSpan = 1;
        foreach ($this->mergedCells as $merge) {
            if (!$cell->isInRange($merge['range'])) {
                continue;
            }

            $rowSpan = 0;
            for ($i = $merge[2]; $i <= $merge[5]; $i++) {
                if (in_array($i, $this->rowSkips)) {
                    continue;
                }
                $rowSpan++;
            }

            if ($merge[0] == $colLetter && $merge[2] == $rowIndex) {
                if ($merge[4] > $colLetterVal) {
                    foreach ($this->tempCols as $c) {
                        $cLetterVal = $this->colLetterVal($c);
                        if ($cLetterVal < $merge[1] || $cLetterVal > $merge[4]) {
                            continue;
                        }
                        if ($cLetterVal < $colLetterVal) {
                            continue;
                        }
                        if ($colLetterVal > $cLetterVal) {
                            break;
                        }
                        $colSpan++;
                    }
                    $colSpan--;
                }
            } else {
                return '';
            }

            if ($rowSpan > 1 || $colSpan > 1) {
                break;
            }
        }
        if ($rowSpan > 1) {
            $return .= ' rowspan="' . $rowSpan . '"';
        }
        if ($colSpan > 1) {
            $return .= ' colspan="' . $colSpan . '"';
        }
        $classes = $this->classCellDefault;

        foreach ($this->classCell as $cl) {
            if ($cell->isInRange($cl[0])) {
                if ($cl[1] == 0) {
                    $classes = $cl[2];
                } else {
                    if ($cl[1] == 1) {
                        foreach ($cl[2] as $class) {
                            $classes[] = $class;
                        }
                    }
                    if ($cl[1] == -1) {
                        foreach ($cl[2] as $class) {
                            $i = array_search($class, $classes);
                            if ($i !== false) {
                                unset($classes[$i]);
                            }
                        }
                    }
                }
            }
        }

        if (count($classes) > 0) {
            $return .= ' class="' . implode(' ', $classes) . '"';
        }


        $return .= '>';
        $value = $cell->getValue() . '';
        foreach ($this->cellValueFormatter as $c) {
            if ($cell->isInRange($c[0])) {
                $value = $c[2]($value);
            }
        }
        if ($this->beautify > 2) {
            $return .= "\n" . $this->intent(4);
        }
        $return .= $value;
        if ($this->beautify > 2) {
            $return .= "\n" . $this->intent(3);
        }
        $return .= '</td>';
        return $return;
    }

    public function setValueFormatter($range)
    {
        $args = func_get_args();
        $priority = 50;
        $clIndex = 1;
        if (count($args) > 1 && is_numeric($args[1])) {
            $priority = $args[1];
            $clIndex++;
        }
        $this->cellValueFormatter[] = [$range, $priority, $args[$clIndex]];
    }

    public function render()
    {
        $return = '';
        if ($this->colTo == null) {
            throw new \Exception('Please set col range end');
        }
        $cols = [];
        $rows = [];

        $firstLetterIndex = array_search($this->colFrom, $this->colLetters);
        $lastLetterIndex = array_search($this->colTo, $this->colLetters);
        for ($i = $firstLetterIndex; $i <= $lastLetterIndex; $i++) {
            if (in_array($this->colLetters[$i], $this->colSkips)) {
                continue;
            }
            $cols[] = $this->colLetters[$i];
        }
        for ($i = $this->rowFrom; $i <= $this->rowTo; $i++) {
            if (in_array($i, $this->rowSkips)) {
                continue;
            }
            $rows[] = $i;
        }
        $this->tempRows = $rows;
        $this->tempCols = $cols;
        foreach ($rows as $rowI) {
            $return .= $this->renderRow($rowI);
        }
        return $return;
    }

    public function __tostring()
    {
        return $this->render();
    }

    private function colLetterVal($letter)
    {
        $return = 0;
        for ($i = 0, $m = strlen($letter); $i < $m; $i++) {
            $return += ord($letter[$i]) + ($i * 256);
        }
        return $return;
    }
}
