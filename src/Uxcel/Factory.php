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

	function __construct() {
		$this->id = uniqid();
		$this->doc = new \DOMDocument();

		$this->supportedExt = array(
			"xlsx"
		);
	}

	public function setSource($source)
	{
		if($source != null) {
			if(strlen(trim($source)) > 0) {
				$this->pathInfo = pathinfo($source);

				$this->fileName = $this->pathInfo["filename"];
				$this->ext = $this->pathInfo["extension"];

				$this->source = $source;
				return $this;
			} else {
				throw new UxcelException("source cannot be empty");
			}
		} else {
			throw new UxcelException("source cannot be null");
		}
	}

	public function setDestination($destination)
	{
		if($destination != null) {
			if(strlen(trim($destination)) > 0) {
				$this->destination = $destination;
				return $this;
			} else {
				throw new UxcelException("destination cannot be empty");
			}
		} else {
			throw new UxcelException("destination cannot be null");
		}
	}

	public function getPathInfo()
	{
		return $this->pathInfo;
	}

	private function fileAvailable()
	{
		if(file_exists($this->source)) {
			return true;
		} else {
			throw new UxcelFileNotFoundException("file {$this->source} doesn't exist");
		}
	}

	private function fileSupported()
	{
		if(in_array($this->ext, $this->supportedExt)) {
			return true;
		}  else {
			throw new UxcelExtensioneNotAllowedException("extension {$this->ext} is doesn't supported");
		}
	}

	public function getExtractDirectory()
	{
		return $this->destination.DIRECTORY_SEPARATOR.$this->fileName.'-'.$this->id;
	}

	public function getWorkBookDirectory()
	{
		$tmp = implode(DIRECTORY_SEPARATOR, [$this->getExtractDirectory(), 'xl']);

		return $tmp.DIRECTORY_SEPARATOR . "workbook.xml";
	}

	public function removeProtectWorkbook()
	{
		$this->doc->load($this->getWorkBookDirectory());

		$elementWbook = $this->doc->documentElement;

		foreach ($elementWbook->getElementsByTagName('workbookProtection') as $childWbook) {
			$elementWbook->removeChild($childWbook);
		}

		$this->doc->save($this->getWorkBookDirectory());

		return $this;
	}

	public function unProtect(){
		if($this->fileAvailable()){
			if($this->fileSupported()){
				$zip = new \ZipArchive();

				if ($zip->open($this->source) === true) {
					$zip->extractTo($this->getExtractDirectory());
						$zip->close();
					} else {
						throw new UxcelException("failed to unzip file");
					}

					$this->removeProtectWorkbook();

					$sheetDir = implode(DIRECTORY_SEPARATOR, [$this->getExtractDirectory(), 'xl', 'worksheets']);

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

					$files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($this->getExtractDirectory()), \RecursiveIteratorIterator::LEAVES_ONLY);

					foreach ($files as $file) {
						if (!$file->isDir()) {
							$file_path = $file->getRealPath();
							$zip->addFile($file_path, substr($file_path, strlen($this->getExtractDirectory()) + 1));
						}
					}

					$zip->close();
					return $this;
				}
			}
		}
	}