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
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Slide\Background\Image;

// Create a new presentation
$ppt = new PhpPresentation();
$ppt->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);

// Create a new slide
$slide = $ppt->getActiveSlide();

// Add an image to the current slide
$shape = $slide->createDrawingShape();
$shape->setName('CodeAdam Logo')
      ->setPath('./images/codeadam.png')
      ->setWidth(100)
      ->setOffsetX(430)
      ->setOffsetY(50);

// Add four text elements to the current slide
$shape = $slide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(180);
$shape->getActiveParagraph()
      ->getAlignment()
      ->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('Adam Thomas');
$textRun->getFont()
      ->setSize(30)
      ->setColor(new Color('FFFFFF'))
      ->setName('Helvetica');

$shape = $slide->createRichTextShape()
      ->setHeight(60)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(240);
$shape->getActiveParagraph()
      ->getAlignment()
      ->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('I Teach Code!');
$textRun->getFont()
      ->setSize(40)
      ->setColor(new Color('eb062c'))
      ->setName('Helvetica');

$shape = $slide->createRichTextShape()
      ->setHeight(40)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(310);
$shape->getActiveParagraph()
      ->getAlignment()
      ->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('Self-taught full-stack developer.');
$textRun->getFont()
      ->setSize(20)
      ->setColor(new Color('FFFFFF'))
      ->setName('Helvetica');

$shape = $slide->createRichTextShape()
      ->setHeight(40)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(360);
$shape->getActiveParagraph()
      ->getAlignment()
      ->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('Learning code and teaching code at Humber Polytechnic, Toronto, Canada.');
$textRun->getFont()
      ->setSize(20)
      ->setColor(new Color('FFFFFF'))
      ->setName('Helvetica');

// Add a background image to the current slide
$background = new Image();
$background->setPath(__DIR__.'/images/ev3.jpg');
$slide->setBackground($background);

// Out put the PPT in memory to a file
$writer = IOFactory::createWriter($ppt, 'PowerPoint2007');
$writer->save(__DIR__ . "/result.pptx");

?>
?>
<h1>Complete</h1>
<p>
      File Exists: <?=file_exists('result.pptx')?>
      <br>
      File Size: <?=filesize('result.pptx')?>
      <br>
      Created: <?=date("F d, Y g:i:s a", filemtime('result.pptx'))?>
</p>