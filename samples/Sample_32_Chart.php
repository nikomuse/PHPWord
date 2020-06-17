<?php
include_once 'Sample_Header.php';

use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Chart as ChartStyle;

// New Word document
echo date('H:i:s'), ' Create new PhpWord object', EOL;

$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->addTitleStyle(1, array('size' => 14, 'bold' => true), array('keepNext' => true, 'spaceBefore' => 240));
$phpWord->addTitleStyle(2, array('size' => 14, 'bold' => true), array('keepNext' => true, 'spaceBefore' => 240));

// 2D charts
$section = $phpWord->addSection();
$section->addTitle(htmlspecialchars('2D charts'), 1);
$section = $phpWord->addSection(array('colsNum' => 1, 'breakType' => 'continuous'));

$sectionStyle = $section->getStyle();
// 21 : 1.5 18 1.5
// 29,7 : 1.7 26.5 1.5
$sectionStyle->setMarginTop(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1.7));
$sectionStyle->setMarginBottom(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1.5));
$sectionStyle->setMarginLeft(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1.5));
$sectionStyle->setMarginRight(\PhpOffice\PhpWord\Shared\Converter::cmToTwip(1.5));

$chartTypes = array('pie', 'doughnut', 'bar', 'column', 'line', 'area', 'scatter', 'radar');
$chartLegend = array('r', 'r', '', 'b', '', '', '', '');
$twoSeries = array('bar', 'column', 'line', 'area', 'scatter', 'radar');
$threeSeries = array('bar', 'line');
$categories = array('A', 'B', 'C', 'D', 'E');
$series1 = array(1, 3, 2, 5, 4);
$series2 = array(3, 1, 7, 2, 6);
$series3 = array(8, 3, 2, 5, 4);
$showGridLines = true;
$showAxisLabels = true;

foreach ($chartTypes as $index => $chartType) {
    $section->addTitle(ucfirst($chartType), 2);
    $chart = $section->addChart($chartType, $categories, $series1);
    $chart->getStyle()->setWidth(Converter::cmToEmu(15))->setHeight(Converter::inchToEmu(2)); // , Converter::inchToEmu(2.5)
    $chart->getStyle()->setShowGridX(false);
    $chart->getStyle()->setShowGridY(true);
    $chart->getStyle()->setMajorTickMark(ChartStyle::TICK_MARK_OUT);
    $chart->getStyle()->setShowAxisLabels($showAxisLabels);
    $chart->getStyle()->addDataLabels(); // ChartStyle::DATA_LABELS_ORIENTATION_RIGHT
    if ($chartType == 'doughnut') {
        $chart->getStyle()->setOption('hole', '55');
    }
    if ($chartLegend[$index]) {
        $chart->getStyle()->addLegend($chartLegend[$index]);
    }
    if (in_array($chartType, $twoSeries)) {
        $chart->addSeries($categories, $series2);
    }
    if (in_array($chartType, $threeSeries)) {
        $chart->addSeries($categories, $series3);
    }
    $section->addTextBreak();
}

// 3D charts
$section = $phpWord->addSection(array('breakType' => 'continuous'));
$section->addTitle(htmlspecialchars('3D charts'), 1);
$section = $phpWord->addSection(array('colsNum' => 2, 'breakType' => 'continuous'));

$chartTypes = array('pie', 'bar', 'column', 'line', 'area');
$multiSeries = array('bar', 'column', 'line', 'area');
$style = array('width' => Converter::cmToEmu(5), 'height' => Converter::cmToEmu(4), '3d' => true, 
                'showAxisLabels' => $showAxisLabels, 'showGridX' => $showGridLines, 'showGridY' => $showGridLines);
foreach ($chartTypes as $index => $chartType) {
    $section->addTitle(ucfirst($chartType), 2);
    $chart = $section->addChart($chartType, $categories, $series1, $style);
    if ($chartLegend[$index]) {
        $chart->getStyle()->addLegend($chartLegend[$index]);
    }
    if (in_array($chartType, $multiSeries)) {
        $chart->addSeries($categories, $series2);
        $chart->addSeries($categories, $series3);
    }
    $section->addTextBreak();
}

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
