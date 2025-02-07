<?php
require($_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_before.php');
require $_SERVER['DOCUMENT_ROOT'] . '/local/scripts/vendor/autoload.php';

use Bitrix\Main\Loader;
use PhpOffice\PhpSpreadsheet\IOFactory;

Loader::includeModule('iblock');

function createIblockIfNotExists($variantProgram) {
    $iblockType = 'polis';
    $iblockCode = 'variant_' . strtolower(str_replace(' ', '_', $variantProgram));

    $existingIblock = CIBlock::GetList([], ['TYPE' => $iblockType, 'CODE' => $iblockCode])->Fetch();
    if ($existingIblock) {
        return $existingIblock['ID'];
    }

    $iblock = new CIBlock;
    $fields = [
        'NAME' => $variantProgram,
        'CODE' => $iblockCode,
        'IBLOCK_TYPE_ID' => $iblockType,
        'SITE_ID' => ['s1'],
        'ACTIVE' => 'Y',
    ];
    return $iblock->Add($fields);
}

function createPropertyIfNotExists($iblockId, $propertyCode, $propertyName) {
    $property = CIBlockProperty::GetList([], ['IBLOCK_ID' => $iblockId, 'CODE' => $propertyCode])->Fetch();
    if ($property) {
        return $property['ID'];
    }

    $ibp = new CIBlockProperty;
    $fields = [
        'IBLOCK_ID' => $iblockId,
        'NAME' => $propertyName,
        'CODE' => $propertyCode,
        'PROPERTY_TYPE' => 'S',
    ];
    return $ibp->Add($fields);
}

function addElementToIblock($iblockId, $fields) {
    $el = new CIBlockElement;
    $el->Add([
        'IBLOCK_ID' => $iblockId,
        'NAME' => $fields['NAME'],
        'ACTIVE_FROM' => $fields['ACTIVE_FROM'],
        'ACTIVE_TO' => $fields['ACTIVE_TO'],
        'PROPERTY_VALUES' => $fields['PROPERTIES'],
    ]);
}

$filePath = $_SERVER['DOCUMENT_ROOT'] . '/local/scripts/act.xlsx';
$spreadsheet = IOFactory::load($filePath);
$worksheet = $spreadsheet->getActiveSheet();
$data = $worksheet->toArray();

$step = 100; // Количество строк за один шаг
$position = isset($_POST['position']) ? (int)$_POST['position'] : 0;
$totalRows = count($data);

if ($position === 0) {
    $position = 1; // Пропуск заголовков
}

for ($i = $position; $i < min($position + $step, $totalRows); $i++) {
    $row = $data[$i];

    $variantProgram = $row[6];
    $iblockId = createIblockIfNotExists($variantProgram);

    // Создаем свойства, если их еще нет
    createPropertyIfNotExists($iblockId, 'DATE_ATT', 'Дата прикрепления');
    createPropertyIfNotExists($iblockId, 'PROGRAMM', 'Программа');
    createPropertyIfNotExists($iblockId, 'AWARD', 'Вознаграждение Исполнителя');
    createPropertyIfNotExists($iblockId, 'CODE', 'Код программы');

    $elementData = [
        'NAME' => $row[1],
        'ACTIVE_FROM' => $row[3],
        'ACTIVE_TO' => $row[4],
        'PROPERTIES' => [
            'DATE_ATT' => $row[2],
            'PROGRAMM' => $row[5],
            'AWARD' => floatval(str_replace([' ', ','], ['', '.'], $row[7])),
            'CODE' => $row[8],
        ],
    ];
    addElementToIblock($iblockId, $elementData);
}

if ($position + $step >= $totalRows) {
    echo json_encode(['status' => 'done']);
} else {
    echo json_encode([
        'status' => 'in_progress',
        'next_position' => $position + $step,
    ]);
}

?>