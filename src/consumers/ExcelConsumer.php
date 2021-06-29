<?php
namespace lucasguo\import\consumers;

use yii\base\BaseObject;
use lucasguo\import\exceptions\InvalidFileException;
use lucasguo\import\components\Importer;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\BaseReader;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Shared\Date;

class ExcelConsumer extends BaseObject implements ConsumerInterface
{
	/**
	 * {@inheritDoc}
	 * @see \backend\modules\import\consumers\ConsumerInterface::consume()
	 * @param $importer Importer
	 */
	public function consume(&$importer)
	{
		try {
			$fileType = IOFactory::identify($importer->file);
			$objReader = IOFactory::createReader($fileType);
            if (!($objReader instanceof Xlsx || $objReader instanceof Xls)) {
                throw new AdapterImportException('File for import is not for this adapter. Actual format: ' . $fileType);
            }
			$objPHPExcel = $objReader->load($importer->file);
		} catch (Exception $e) {
			throw new InvalidFileException();
		}
		$sheet = $objPHPExcel->getSheet(0);
		$highestRow = $sheet->getHighestRow();
		$data = [];
		$highestCol = $this->getNameFromNumber($importer->getMaxColIndex() + 1);
		for ($i = $importer->skipRowsCount + 1; $i <= $highestRow; $i++) {
			$rowDataArray = $sheet->rangeToArray('A' . $i . ':' . $highestCol . $i);
			$rowData = $rowDataArray[0];
			$skip = false;
			foreach ($importer->getRequiredCols() as $col) {
				if ($rowData[$col] == null) {
					$skip = true;
					break;
				}
			}
			if ($skip) {
				$importer->addSkipRow($i);
			} else {
				$data[$i] = $rowData;
			}
		}
		unset($objPHPExcel);
		unset($objReader);
		return $data;
	}
	
	protected function getNameFromNumber($num) 
	{
	    $numeric = ($num - 1) % 26;
	    $letter = chr(65 + $numeric);
	    $num2 = intval(($num - 1) / 26);
	    if ($num2 > 0) {
	        return getNameFromNumber($num2) . $letter;
	    } else {
	        return $letter;
	    }
	}
	
	
}
