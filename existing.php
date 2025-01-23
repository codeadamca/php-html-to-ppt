<?php

require __DIR__ . '/vendor/autoload.php';

/*
 *
 * This example makes a 16:9 presentation which converts to 
 * 960 pixels wide by 540 pixels high. 
 * 
 * This example adds a locel image, text, and a background
 * image.
 *
 */

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Slide\Background\Image;

// Create a new presentation
$ppt = new PhpPresentation();
$ppt->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);
$ppt->removeSlideByIndex(0);

// Open an existing PPT file and add one slide
$reader = IOFactory::createReader('PowerPoint2007');
$slides = $reader->load(__DIR__ . '/sample.pptx');
$slides = $slides->getAllSlides();
$ppt->addExternalSlide($slides[1], 0);

/*
 *
 * For some reason if the code attempts to add two slides from an 
 * external presentation, it causes the resulting PPT file to be
 * corrupt. But if you re-open the existing PPT file in between adds
 * the resulting PPT file is fine. 
 * 
 */

// Reopen an existing PPT file and add one slide
$reader = IOFactory::createReader('PowerPoint2007');
$slides = $reader->load(__DIR__ . '/sample.pptx');
$slides = $slides->getAllSlides();
$ppt->addExternalSlide($slides[2], 1);

// Out put the PPT in memory to a file
$writer = IOFactory::createWriter($ppt, 'PowerPoint2007');
$writer->save(__DIR__ . "/result.pptx");

?>
<h1>Complete</h1>
<p>
      File Exists: <?=file_exists('result.pptx')?>
      <br>
      File Size: <?=filesize('result.pptx')?>
      <br>
      Created: <?=date("F d, Y g:i:s a", filemtime('result.pptx'))?>
</p>