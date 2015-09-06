<?php

/*******************************************************
* This class was coded for Matthew1471 by Andrew Gillard
* -- www.lorddeath.net --
* All copyrights belong to him
********************************************************/ 

class Thumbnailer {
	var $Filename;
	var $MaxX;
	var $MaxY;
	var $CurrentX;
	var $CurrentY;
	var $NewX;
	var $NewY;
	var $Format;
	var $JPEGQuality = 75;
	var $SaveFilename = null;

	function CalcSizes() {
		if ($this->CurrentY > $this->MaxY || $this->CurrentX > $this->MaxX) {
			//Image is too big
			if (($this->CurrentX/$this->CurrentY) > ($this->MaxX/$this->MaxY)) {
				//Needs width to be decreased
				$this->NewX = $this->MaxX;
				$this->NewY = round(($this->CurrentY/$this->CurrentX) * $this->MaxX);
			} else if ($this->CurrentX == $this->CurrentY) {
				//Square
				$max = ($this->MaxX > $this->MaxY) ? $this->MaxY : $this->MaxX;
				$this->NewX = $max;
				$this->NewY = $max;
			} else {
				//Needs height to be decreased
				$this->NewY = $this->MaxY;
				$this->NewX = round(($this->CurrentX/$this->CurrentY) * $this->MaxY);
			}
		} else {
			$this->NewX = $this->CurrentX;
			$this->NewY = $this->CurrentY;
		}
	}

	function Generate() {
		if (strpos($this->Filename, "..") !== false ||
			substr($this->Filename, 1, 1) == ":" ||
			substr($this->Filename, 0, 1) == "/")
			die("Invalid filename -- cannot access files below this directory.");
		if (!$ImageInfo = @getimagesize($this->Filename))
			die("Invalid image.");
		else {
			switch($ImageInfo[2]) {
				case 1: //GIF
					$image = imagecreatefromgif($this->Filename);
					break;
				case 2: //JPEG
					$image = imagecreatefromjpeg($this->Filename);
					break;
				case 3: //PNG
					$image = imagecreatefrompng($this->Filename);
					break;
				default:
					die("Invalid image type. Valid formats are: GIF, JPEG, PNG.");
			}
			$this->CurrentX = $ImageInfo[0];
			$this->CurrentY = $ImageInfo[1];
			$this->CalcSizes();
			$newimg = imagecreatetruecolor($this->NewX, $this->NewY);
			imagecopyresampled($newimg, $image, 0, 0, 0, 0, $this->NewX, $this->NewY, $this->CurrentX, $this->CurrentY);
			switch(strtolower($this->Format)) {
				case "gif":
					header("Content-type: image/gif");
					imagegif($newimg, $this->SaveFilename);
					break;
				case "jpg":
				case "jpeg":
					header("Content-type: image/jpeg");
					imagejpeg($newimg, $this->SaveFilename, $this->JPEGQuality);
					break;
				case "png":
					header("Content-type: image/png");
					imagepng($newimg, $this->SaveFilename);
					break;
				default:
					die("Invalid output format. Valid formats are: GIF, JPEG, PNG.");
			}
			imagedestroy($image);
			imagedestroy($newimg);
		}
	}
};

$Thumbnailer = new Thumbnailer();

$Thumbnailer->Filename = $_GET['f'];
$Thumbnailer->MaxX = 600;
$Thumbnailer->MaxY = 90;
$Thumbnailer->Format = "JPEG";
$Thumbnailer->JPEGQuality = 50;

/* Matthew tries to learn PHP */
if (strpos($Thumbnailer->Filename, '/') === false) {
 $Thumbnailer->SaveFilename = 'Thumbnails/tn' . $Thumbnailer->Filename;
} else {
 $Thumbnailer->SaveFilename = preg_replace('£^(.*)/([^/]*)$£', '\\1/Thumbnails/tn\\2', $Thumbnailer->Filename);
}

 $Thumbnailer->Generate();
 readfile($Thumbnailer->SaveFilename);

//var_dump($Thumbnailer);
?>