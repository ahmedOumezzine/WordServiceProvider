<?php
namespace SilexCasts\Provider\Word;
use SilexCasts\Provider\Word\PHPWord_Template;
class PHPWord
{
    public function Hello(){
         return "hello";
    }
    public function loadTemplate($strFilename) {
        if(file_exists($strFilename)) {
            $document = new PHPWord_Template($strFilename);

            $document->setValue('Value1', "eszereqsdqsdzr");
            $document->setValue('Value2', 'Mercury');
            $document->setValue('Value3', 'Venus');
            $document->setValue('Value4', 'Earth');
            $document->setValue('Value5', 'Mars');
            $document->setValue('Value6', 'Jupiter');
            $document->setValue('Value7', 'Saturn');
            $document->setValue('Value8', 'Uranus');
            $document->setValue('Value9', 'Neptun');
            $document->setValue('Value10', 'Pluto');
            $document->setValue('myReplacedValue', 'zzzzzzzzzzzzzzzzzzzzzzzzzzz');

            $document->setValue('weekday', date('l'));
            $document->setValue('time', date('H:i'));

            $document->save('word/Solarsystem.docx');

            return "er";
        } else {
           return "not fond";
        }
    }
}

?>