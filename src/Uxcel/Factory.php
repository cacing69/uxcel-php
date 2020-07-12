<?php
namespace Uxcel;

use Uxcel\Exceptions\UxcelException;
use Uxcel\Exceptions\UxcelFileNotFoundException;
use Uxcel\Exceptions\UxcelExtensioneNotAllowedException;

class Factory {
	private $supportedExt;
	public $fileName;
	public $id;
	public $ext;
	private $target;
	private $dir;
	private $doc;
	private $pathInfo;
	private $log;

	function __construct() {
		$this->id = uniqid();
		$this->doc = new \DOMDocument();
		$this->supportedExt = array(
			"xlsx"
		);
		$this->log = array();
	}

	public function setSource($source)
	{
		$this->pathInfo = pathinfo($source);

		$this->fileName = $this->pathInfo["filename"];
		$this->ext = $this->pathInfo["extension"];

		$this->source = $source;
		return $this;
	}

	public function setDestination($destination)
	{
		$this->destination = $destination;
		return $this;
	}

	public function getPathInfo()
	{
		return $this->pathInfo;
	}

	public function unProtect(){
		if(file_exists($this->source)) {
			if(in_array($this->ext, $this->supportedExt)){
				$zip = new \ZipArchive();

				$dir = $this->destination.DIRECTORY_SEPARATOR.$this->fileName.'-'.$this->id;

				if ($zip->open($this->source) === true) {
					$zip->extractTo($dir);
					$zip->close();
				} else {
					throw new UxcelException("Uxcel : failed to unzip file");
				}

				// remove protect workbook first
				$workbookDir = implode(DIRECTORY_SEPARATOR, [$dir, 'xl']);

				$this->doc->load($workbookDir. DIRECTORY_SEPARATOR . "workbook.xml");

				$elementWbook = $this->doc->documentElement;

				foreach ($elementWbook->getElementsByTagName('workbookProtection') as $childWbook) {
					$elementWbook->removeChild($childWbook);
				}

				$sheetDir = implode(DIRECTORY_SEPARATOR, [$dir, 'xl', 'worksheets']);

				foreach (glob($sheetDir . DIRECTORY_SEPARATOR . '*.xml') as $sheet) {
					$this->doc->load($sheet);
					$elementSheet = $this->doc->documentElement;

					foreach ($elementSheet->getElementsByTagName('sheetProtection') as $child) {
						$elementSheet->removeChild($child);
					}

					$this->doc->save($sheet);
				}

				$zip = new \ZipArchive();
				$zip->open($this->destination.DIRECTORY_SEPARATOR.$this->fileName."-".$this->id.".".$this->ext, \ZIPARCHIVE::CREATE);

				$files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($dir), \RecursiveIteratorIterator::LEAVES_ONLY);

				foreach ($files as $file) {
					if (!$file->isDir()) {
						$file_path = $file->getRealPath();
						$zip->addFile($file_path, substr($file_path, strlen($dir) + 1));
					}
				}

				$zip->close();
				return $this;
			} else {
				throw new UxcelExtensioneNotAllowedException("extension {$this->ext} is doesn't supported");
			}
		} else {
			throw new UxcelFileNotFoundException("file {$this->source} doesn't exist");
		}
	}
}