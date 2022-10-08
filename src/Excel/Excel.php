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
     * @var array Security mode for Microsoft Excel in open files by prog
     */
    public static $msoAutomationSecurity = [
        'msoAutomationSecurityLow' => 1,
        'msoAutomationSecurityByUi' => 2,
        'msoAutomationSecurityForceDisable' => 3
    ];

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

    /**
     * Adding a custom list for incremental copy and/or custom sort
     * @param $listArray
     * @param bool $byRow
     * @return mixed
     */
    public function addCustomlList($listArray, bool $byRow = false): Excel
    {
        $this->com->AddCustomList($listArray, $byRow);

        return $this;
    }

    /**
     * Return a collection of AddIns, represent all add-ins
     * @return AddIns
     * @noinspection PhpIncompatibleReturnTypeInspection
     * @noinspection PhpDocSignatureInspection
     * @noinspection PhpUndefinedClassInspection
     */
    public function getAddIns()
    {
        return $this->com->AddIns;
    }

    /**
     * Return a collection of AddIns2, represent all modules actually open in Microsoft Excel. Installed or not.
     * @return AddIns2
     * @noinspection PhpIncompatibleReturnTypeInspection
     * @noinspection PhpDocSignatureInspection
     * @noinspection PhpUndefinedClassInspection
     */
    public function getAddIns2()
    {
        return $this->com->AddIns2;
    }

    /**
     * Set the display the message before replace cells with data over change with drag & drop
     * @param bool $alert
     * @return Excel
     */
    public function setAlertBeforeOverwrting(bool $alert): Excel
    {
        $this->com->AlertBeforeOverwriting = $alert;

        return $this;
    }

    /**
     * Get the display the message before replace cells with data over change with drag & drop
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAlertBeforeOverwriting(): bool
    {
        return $this->com->AlertBeforeOverwriting;
    }

    /**
     * Set the name of alternative directory of startup
     * @param string $path
     * @return Excel
     */
    public function setAltStartupPath(string $path): Excel
    {
        $this->com->AltStartupPath = mb_convert_encoding($path, $this->charset);

        return $this;
    }

    /**
     * Get the name of alternative directory of startup
     * @return string
     */
    public function getAltStartupPath(): string
    {
        return $this->com->AltStartupPath;
    }

    /**
     * Set the ClearType value for display the fonts in menu, ribbon & dialog
     * @param bool $clearType
     * @return Excel
     */
    public function setAlwaysUseClearType(bool $clearType): Excel
    {
        $this->com->AlwaysUseClearType = $clearType;

        return $this;
    }

    /**
     * Get the ClearType value for display the fonts in menu, ribbon & dialog
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAlwaysUseClearType(): bool
    {
        return $this->com->AlwaysUseClearType;
    }

    /**
     * Get if the XML functionalities of Microsoft Excel is available
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getArbitraryXMLSupportAvailable(): bool
    {
        return $this->com->ArbitraryXMLSupportAvailable;
    }

    /**
     * Set the Microsoft Excel ask user to update a links in the open file used
     * @param bool $ask
     * @return Excel
     */
    public function setAskToUpdateLinks(bool $ask): Excel
    {
        $this->com->AskToUpdateLinks = $ask;

        return $this;
    }

    /**
     * Set the Microsoft Excel ask user to update a links in the open file used
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAskToUpdateLinks(): bool
    {
        return $this->com->AskToUpdateLinks;
    }

    /**
     * Get the Microsoft Help Office viewer of Microsoft Excel
     * @return IAssistance
     * @noinspection PhpUndefinedClassInspection
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAssistance(): IAssistance
    {
        return $this->com->Assistance;
    }

    /**
     * Get the AutoCorrect object of represent the Decorrect auto function of Microsoft Excel
     * @return AutoCorrect
     * @noinspection PhpUndefinedClassInspection
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAutoCorrect(): AutoCorrect
    {
        return $this->com->AutoCorrect;
    }

    /**
     * Set the Microsoft Excel auto formatting the hyperlinks
     * @param bool $autoFormatting
     * @return Excel
     */
    public function setAutoFormatAsYouTypeReplaceHyperlinks(bool $autoFormatting): Excel
    {
        $this->com->AutoFormatAsYouTypeReplaceHyperlinks = $autoFormatting;

        return $this;
    }

    /**
     * Get the Microsoft Excel auto formatting the hyperlinks
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAutoFormatAsYouTypeReplaceHyperlinks(): bool
    {
        return $this->com->AutoFormatYouTypeReplaceHyperlinks;
    }

    /**
     * Set the msoAutomationSecurity const for represent the security mode used by Microsoft Excel in open files by prog
     * @param int $msoAutomationSecurity
     * @return Excel
     * @throws Exception
     */
    public function setAutomationSecurity(int $msoAutomationSecurity): Excel
    {
        if (
            $msoAutomationSecurity < self::$msoAutomationSecurity['msoAutomationSecurityLow']
            || $msoAutomationSecurity > self::$msoAutomationSecurity['msoAutomationSecurityForceDisable']
        ) {
            throw new Exception("Please used the msoAutomationSecurity static property, for used the valid value");
        }

        $this->com->AutomationSecurity = $msoAutomationSecurity;

        return $this;
    }

    /**
     * Get the msoAutomationSecurity const for represent the security mode used by Microsoft Excel in open files by prog
     * @return int
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAutomationSecurity(): int
    {
        return $this->com->AutomationSecurity;
    }

    /**
     * Set the auto apply the multiplication by 100 for the formatting cells in percentage
     * @param bool $percentEntry
     * @return Excel
     */
    public function setAutoPercentEntry(bool $percentEntry): Excel
    {
        $this->com->AutoPercentEntry = $percentEntry;

        return $this;
    }

    /**
     * Get the auto apply the multiplication by 100 for the formatting cells in percentage
     * @return bool
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAutoPercentEntry(): bool
    {
        return $this->com->AutoPercentEntry;
    }

    /**
     * Get the AutoRecover object for get the files format in the time interval
     * @return AutoRecover
     * @noinspection PhpUndefinedClassInspection
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getAutoRecover(): AutoRecover
    {
        return $this->com->AutoRecover;
    }

    /**
     * Get the build number of Microsoft Excel
     * @return int
     * @noinspection PhpIncompatibleReturnTypeInspection
     */
    public function getBuild(): int
    {
        return $this->com->Build;
    }
}
