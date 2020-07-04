<?php
namespace Uxcel;

class Factory {
	private $source;
	public $filename;
	public $id;
	public $ext;
	private $target;
	private $dir;
	private $doc;
	private $pathInfo;

	function __construct() {
		$this->id = uniqid();
        $this->doc = new \DOMDocument();
    }

	public function setSource($source)
	{
		$this->pathInfo = pathinfo($source);

		$this->filename = $this->pathInfo["filename"];
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
		$zip = new \ZipArchive();

		$dir = $this->destination.DIRECTORY_SEPARATOR.$this->filename.'-'.$this->id;

		if ($zip->open($this->source) === true) {
			$zip->extractTo($dir);
			$zip->close();
		} else {
			throw new \Exception("Uxcel : failed to unzip file", 1);
		}

		$sheetDir = implode(DIRECTORY_SEPARATOR, [$dir, 'xl', 'worksheets']);
		foreach (glob($sheetDir . DIRECTORY_SEPARATOR . '*.xml') as $sheet) {
			$this->doc->load($sheet);
			$element = $this->doc->documentElement;

			foreach ($element->getElementsByTagName('sheetProtection') as $child) {
				$element->removeChild($child);
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
	}
}