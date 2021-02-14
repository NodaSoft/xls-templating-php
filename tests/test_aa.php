<?php

use NodaSoft\PhpXlsTemplating\ProcessOptions;
use NodaSoft\PhpXlsTemplating\SheetTplData;
use NodaSoft\PhpXlsTemplating\Templating;

include_once 'vendor/autoload.php';

$t = new Templating();
$t->setTplFileName('tests/test.xlsx');
$t->setResultFileName('tests/test.result.xlsx');
$options = new ProcessOptions();
$options->copyStyles = true;
$options->insertInsteadOfCopy = true;
$t->setOptions($options);
$t->run([
    new SheetTplData('', [
    'testVar' => 'tessstt',
    'image1' => 'tests/image.jpg',
    'cond' => true,
    'condPos' => 'condPos1',
    'condNeg' => 'condNeg1',
    'cond2' => false,
    'cond2Pos' => 'pos2',
    'cond2Neg' => 'neg2',
    'positions1' => [
        [
            'a' => 'a1',
            'b' => 'b1',
            'c' => 'c1',
            'd' => 'd1',
        ],
        [
            'a' => 'a2',
            'b' => 'b2',
            'c' => 'c2',
            'd' => 'd2',
        ],
        [
            'a' => 'a3',
            'b' => 'b3',
            'c' => 'c3',
            'd' => 'd3',
        ],
    ],
    'positions2' => [
        [
            'a' => 'a1',
            'b' => 'b1',
            'c' => 'c1',
            'd' => 'd1',
        ],
        [
            'a' => 'a2',
            'b' => 'b2',
            'c' => 'c2',
            'd' => 'd2',
        ],
        [
            'a' => 'a3',
            'b' => 'b3',
            'c' => 'c3',
            'd' => 'd3',
        ],
        [
            'a' => 'a4',
            'b' => 'b4',
            'c' => 'c4',
            'd' => 'd4',
        ],
    ],

])]); // Mpdf