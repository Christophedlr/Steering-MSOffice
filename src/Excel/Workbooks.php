<?php


namespace Christophedlr\StMso\Excel;


use COM;
use com_exception;

/**
 * The Workbooks class for manage the Workbooks Excel with the Windows COM Interface
 * @package Christophedlr\StMso\Excel
 * @license MIT
 * @copyright Christophe DALOZ - DE LOS RIOS, 2022
 */
class Workbooks
{
    private $com;
    private $parent;
    private $charset;

    public static $xlWBATemplate = [
        'xlWBATChart' => -4109,
        'xlWBATExcel4IntlMacroSheet' => 4,
        'xlWBATExcel4MacroSheet' => 3,
        'xlWBATWorksheet' => -4167
    ];

    public function __construct(COM $com, Excel $parent, string $charset = "windows-1252")
    {
        $this->com = $com->Workbooks;
        $this->parent = $parent;
        $this->charset = $charset;
    }

    public function application()
    {
        return $this->com->Application;
    }

    /**
     * Get the number of Workbooks are opens
     * @return int
     */
    public function getCount(): int
    {
        return $this->com->Count;
    }

    /**
     * Get the Workbook
     * @param $item
     * @return Workbook
     * @noinspection PhpUndefinedMethodInspection
     * @noinspection PhpUndefinedClassInspection
     * @throws Exception
     */
    public function getItem($item)
    {
        if (is_int($item)) {
            return $this->com->Item($item);
        } else if (is_string($item)) {
            return $this->com->Item(mb_convert_encoding($item, $this->charset));
        }

        throw new Exception("Item value is integer or string only");
    }

    /**
     * Get the Excel parent application
     * @return Excel
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Add a new Workshhet
     * @param XlWBATemplate|string $template
     * @noinspection PhpUndefinedClassInspection
     * @noinspection PhpUndefinedMethodInspection
     * @noinspection PhpDocSignatureInspection
     * @return Worksheet
     */
    public function add($template = -4167)
    {
        if (is_string($template)) {
            return $this->com->Add(mb_convert_encoding($template, $this->charset));
        } else if (is_int($template)) {
            return $this->com->Add($template);
        }

        return $this->com->Add;
    }

    /**
     * Verify if possible to extract the Workbook of file
     * @param string $filename
     * @return bool
     * @noinspection PhpUndefinedMethodInspection
     */
    public function canCheckOut(string $filename): bool
    {
        return $this->com->CanCheckOut(mb_convert_encoding($filename, $this->charset));
    }

    /**
     * Extract Workbook of specified file
     * @param string $filename
     * @return string
     * @noinspection PhpUndefinedMethodInspection
     */
    public function checkOut(string $filename): string
    {
        return $this->com->CheckOut(mb_convert_encoding($filename, $this->charset));
    }

    /**
     * Close the Workbooks
     */
    public function close()
    {
        $this->com->Close;
    }
}
