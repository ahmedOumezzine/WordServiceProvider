<?php
/**
 * Created by PhpStorm.
 * User: AhmedOumezzine
 * Date: 6/8/2017
 * Time: 9:29 AM
 */
namespace SilexCasts\Provider\Word;

class  PHPWord_Template
{
    /**
     * ZipArchive
     *
     * @var ZipArchive
     */
    private $_objZip;

    /**
     * Temporary Filename
     *
     * @var string
     */
    private $_tempFileName;

    /**
     * Document XML
     *
     * @var string
     */
    private $_documentXML;


    /**
     * Create a new Template Object
     *
     * @param string $strFilename
     */
    public function __construct($strFilename)
    {
        $path = dirname($strFilename);
        $this->_tempFileName = $path . DIRECTORY_SEPARATOR . time() . '.docx';
        copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

        // Create the Object.
        $this->_objZip = new \ZipArchive;

        $token1 = 'Hello World!';
        $token2 = 'Your mother smelt of elderberries, and your father was a hamster!';
        if ( $this->_objZip->open($this->_tempFileName, \ZipArchive::CREATE) !== TRUE) {
            echo "Cannot open $this->_tempFileName :( ";
            die;
        }
         $this->_documentXML =  $this->_objZip->getFromName('word/document.xml');
        /* $this->_documentXML = str_replace('{Value4}', $token1,  $this->_documentXML);
         $this->_documentXML = str_replace('{Value5}', $token2,  $this->_documentXML);
        if ( $this->_objZip->addFromString('word/document.xml',  $this->_documentXML)) {
            echo 'File written!';
        } else {
            echo 'File not written.  Go back and add write permissions to this folder!l';
        }
         $this->_objZip->close();

        var_dump( $this->_documentXML);
    */
        }

    /**
     * Set a Template value
     *
     * @param mixed $search
     * @param mixed $replace
     */
    public function setValue($search, $replace)
    {
        if (substr($search, 0, 2) !== '${' && substr($search, -1) !== '}') {
            $search = '${' . $search . '}';
        }

        if (!is_array($replace)) {
            $replace = utf8_encode($replace);
        }

        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);

    }

    /**
     * Save Template
     *
     * @param string $strFilename
     */
    public function save($strFilename)
    {
        if (file_exists($strFilename)) {
            unlink($strFilename);
        }
        if ( $this->_objZip->addFromString('word/document.xml',  $this->_documentXML)) {
            echo 'File written!';
        } else {
            echo 'File not written.  Go back and add write permissions to this folder!l';
        }

        // Close zip file
        if (  $this->_objZip->close() === false) {
            throw new \Exception('Could not close zip file.');
        }

        rename($this->_tempFileName, $strFilename);
    }
}