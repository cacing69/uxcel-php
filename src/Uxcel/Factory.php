<?php
namespace Uxcel;

class Factory {
	public $source;
	protected $target;
	public $dir;
	public $id;
	private $ext;
	private $doc;
	private $pathInfo;

	function __construct() {
        $this->doc = new \DOMDocument();
    }

	protected function setSource($source)
	{
		$this->pathInfo = pathinfo($source);
		$this->source = $source;
	}

	protected function setTarget($target)
	{
		$this->target = $target;
	}

	function unProtect($source, $target){



		$zip = new \ZipArchive();

		$dir = $this->target.DIRECTORY_SEPARATOR.$this->pathInfo["filename"].'-'.$this->id ;

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
		
		$zip->open($this->target.DIRECTORY_SEPARATOR.$this->pathInfo["filename"]."-".$this->id.".".$this->pathInfo["extension"], \ZIPARCHIVE::CREATE);

		$files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($dir), \RecursiveIteratorIterator::LEAVES_ONLY);

		foreach ($files as $file) {
			if (!$file->isDir()) {
				$file_path = $file->getRealPath();
				$zip->addFile($file_path, substr($file_path, strlen($dir) + 1));
			}
		}

		$zip->close();
	}
}