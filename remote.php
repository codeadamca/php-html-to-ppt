<?php

require __DIR__ . '/vendor/autoload.php';

/*
 *
 * This example makes a 16:9 presentation which converts to 
 * 960 pixels wide by 540 pixels high. 
 * 
 * This example adds an image using a URL.
 *
 */

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;

$image = 'https://console.codeadam.ca/storage/articles/mywzrQxxvBHS4qDb71CL0EEtf0bcxJ6mtupNz4In.png';
$contents = file_get_contents($image);
$base64 = base64_encode($contents);
$image = 'data:image/png;base64,' . $base64;

// Create a new presentation
$ppt = new PhpPresentation();
$ppt->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);

// Create a new slide
$slide = $ppt->getActiveSlide();

// Add an image to the current slide
$shape = new Base64();
$shape->setName('Sample image')
      ->setData($image)
      ->setWidth(400)
      ->setHeight(400)
      ->setOffsetX(10)
      ->setOffsetY(10);
$slide->addShape($shape);

// Out put the PPT in memory to a file
$writer = IOFactory::createWriter($ppt, 'PowerPoint2007');
$writer->save(__DIR__ . "/sample.pptx");

?>
<h1>Complete</h1>
<p>
      File Exists: <?=file_exists('sample.pptx')?>
      <br>
      File Size: <?=filesize('sample.pptx')?>
      <br>
      Created: <?=date("F d, Y g:i:s a", filemtime('sample.pptx'))?>
</p>