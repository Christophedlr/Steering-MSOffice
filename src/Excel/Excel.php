<?php


namespace Christophedlr\StMso\Excel;


use COM;
use com_exception;

/**
 * The Excel class permit you create an instance of Excel application with the Windows COM Interface
 * @package Christophedlr\StMso\Excel
 * @license MIT
 * @copyright Christophe DALOZ - DE LOS RIOS, 2022
 */
class Excel
{
    /**
     * @var COM COM Interface instance
     * By default, display the Excel application
     */
    private $com;

    /**
     * @var string Charset used in the Excel application
     */
    private $charset;

    public function __construct(string $charset = "windows-1252")
    {
        try {
            $this->com = new COM("excel.application");
        } catch (com_exception $e) {
            echo 'Unable to instantiate a Microsoft Excel';
        }

        $this->visible(true);
    }

    /**
     * get the COM instance of Excel application
     * @return COM Excel application (COM Interface)
     */
    public function application(): COM
    {
        return $this->com;
    }

    /**
     * @param bool $visible Enable or disable the display or Excel application
     * @return Excel Instance of Excel application (fluent)
     */
    public function visible(bool $visible): Excel
    {
        $this->com->Visible = $visible;

        return $this;
    }

    /**
     * Close the Excel application
     */
    public function close()
    {
        $this->com->Quit();
    }

    /**
     * Show the Excel alert message
     * @param bool $alert true if Microsoft Excel display alert message
     * @return Excel Instance of Excel application (fluent)
     */
    public function displayAlerts(bool $alert): Excel
    {
        $this->com->DisplayAlerts($alert);

        return $this;
    }

    /**
     * The active cell of active window
     * @return Range Return the Range instance
     * @throws com_exception
     * @todo Create Range class
     */
    public function getActiveCell()
    {
        $return = $this->com->ActiveCell;

        if (is_null($return)) {
            throw new com_exception("No active sheet find");
        }

        return $return;
    }

    /**
     * The active chart of active window
     * @return Chart Return the Chart instance
     * @throws Exception
     * @todo Create Chart class
     */
    public function getActiveChart()
    {
        try {
            $return = $this->com->ActiveChart;
        } catch (com_exception $e) {
            throw new Exception("No active workbook find", 91);
        } finally {
            $this->close();
        }

        if (is_null($return)) {
            throw new Exception("No active chart find", 91);
        }

        return $return;
    }

    /**
     * Return a long for represent the encrypted session associate to active document
     * @return int long representation of encrypted session
     */
    public function getActiveEncryptionSession(): int
    {
        return $this->com->ActiveEncryptionSession;
    }

    /**
     * Change the name of active printer
     * @param string $name Name of printer
     * @return Excel Instance of Excel application (fluent)
     */
    public function setActivePrinter(string $name): Excel
    {
        $this->com->ActivePrinter = mb_convert_encoding($name, $this->charset);

        return $this;
    }

    /**
     * Get the name of active printer
     * @return string Name of printer
     */
    public function getActivePrinter(): string
    {
        return $this->com->ActivePrinter;
    }

    /**
     * Get the active window of protected display
     * @return ProtectedViewWindow Object represent protected window
     * @throws Exception
     * @todo Create ProtecedViewWindow class
     */
    public function getActiveProtectedViewWindow()
    {
        $return = $this->com->ActiveProtectedViewWindow;

        if (is_null($return)) {
            throw new Exception("No active window in protected mode";
        }

        return $return;
    }

    /**
     * Get the active sheet
     * @return Worksheet Active sheet
     * @throws Exception
     * @todo Create Worksheet class
     */
    public function getActiveSheet()
    {
        $return = $this->com->ActiveSheet;

        if (is_null($return)) {
            throw new Exception("No active sheet find";
        }

        return $return;
    }

    /**
     * Get the active window
     * @return Window Active window
     * @throws Exception
     * @todo Create Window class
     */
    public function getActiveWindow()
    {
        $return = $this->com->ActiveWindow;

        if (is_null($return)) {
            throw new Exception("No active window find";
        }

        return $return;
    }

    /**
     * Get the active workbook
     * @return Workbook Active workbook
     * @throws Exception
     */
    public function getActiveWorkbook()
    {
        $return = $this->com->ActiveWorkbook;

        if (is_null($return)) {
            throw new Exception("No active workbook find";
        }

        return $return;
    }
}
