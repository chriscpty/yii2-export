<?php

/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @version   2.0.0
 */

namespace kartik\export;

use Closure;
use Exception;
use kartik\base\TranslationTrait;
use kartik\dialog\Dialog;
use kartik\dynagrid\Dynagrid;
use kartik\grid\GridView;
use OpenSpout\Common\Entity\Cell as OpenspoutCell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Writer\AutoFilter;
use OpenSpout\Writer\Common\AbstractOptions;
use OpenSpout\Writer\Common\Entity\Sheet;
use OpenSpout\Writer\CSV\Options as OpenspoutCsvOptions;
use OpenSpout\Writer\CSV\Writer as OpenspoutCsvWriter;
use OpenSpout\Writer\Exception\Border\InvalidNameException;
use OpenSpout\Writer\Exception\Border\InvalidStyleException;
use OpenSpout\Writer\Exception\Border\InvalidWidthException;
use OpenSpout\Writer\Exception\InvalidSheetNameException;
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\ODS\Options as OpenspoutOdsOptions;
use OpenSpout\Writer\ODS\Writer as OpenspoutOdsWriter;
use OpenSpout\Writer\XLSX\Entity\SheetView;
use OpenSpout\Writer\XLSX\Options;
use OpenSpout\Writer\XLSX\Properties;
use OpenSpout\Writer\XLSX\Writer;
use Throwable;
use Yii;
use yii\base\InvalidConfigException;
use yii\base\Model;
use yii\base\Widget;
use yii\data\ActiveDataProvider;
use yii\data\ArrayDataProvider;
use yii\data\BaseDataProvider;
use yii\db\ActiveQueryInterface;
use yii\db\QueryInterface;
use yii\grid\ActionColumn;
use yii\grid\Column;
use yii\grid\DataColumn;
use yii\grid\SerialColumn;
use yii\helpers\ArrayHelper;
use yii\helpers\Html;
use yii\helpers\Inflector;
use yii\helpers\Json;
use yii\helpers\Url;
use yii\web\JsExpression;
use yii\web\Request;
use yii\web\View;

/**
 * Export menu widget. Export tabular data to various formats using the openspout library
 * by reading data from a dataProvider - with configuration very similar to a GridView.
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since  1.0
 */
class ExportMenu extends GridView
{
    use TranslationTrait;

    /**
     * @var string UTF-8 encoding
     */
    public const ENCODING_UTF8 = 'utf-8';
    /**
     * @var string CSV (comma separated values) export format
     */
    public const FORMAT_CSV = 'Csv';
    /**
     * @var string Text export format
     */
    public const FORMAT_TEXT = 'Txt';
    /**
     * @var string Microsoft Excel 2007+ export format
     */
    public const FORMAT_EXCEL_X = 'Xlsx';
    /**
     * @var string OpenOffice document format
     */
    public const FORMAT_ODS = 'Ods';
    /**
     * @var string Set download target for grid export to a popup browser window
     */
    public const TARGET_POPUP = '_popup';
    /**
     * @var string Set download target for grid export to a dynamically generated iframe
     */
    public const TARGET_IFRAME = '_iframe';
    /**
     * @var string Set download target for grid export to the same open document on the browser
     */
    public const TARGET_SELF = '_self';
    /**
     * @var string Set download target for grid export to a new window that auto closes after download
     */
    public const TARGET_BLANK = '_blank';

    /**
     * @var string the target for submitting the export form, which will trigger the download of the exported file.
     * Must be one of the `TARGET_` constants. Defaults to [[TARGET_SELF]]. Note if you set [[stream]] to `false`,
     * then this will be always overridden to [[TARGET_SELF]].
     */
    public string $target = self::TARGET_SELF;

    /**
     * @var array configuration settings for the Krajee dialog widget that will be used to render alerts and
     * confirmation dialog prompts
     * @see http://demos.krajee.com/dialog
     */
    public $krajeeDialogSettings = [];

    /**
     * @var bool whether to show a confirmation alert dialog before download. This confirmation dialog will notify
     * user about the type of exported file for download and to disable popup blockers. Defaults to `true`.
     */
    public bool $showConfirmAlert = true;

    /**
     * @var bool whether to enable the yii gridview formatter component. Defaults to `true`. If set to `false`, this
     * will render content as `raw` format.
     */
    public bool $enableFormatter = true;

    /**
     * @var bool whether to render the export menu as bootstrap button dropdown widget. Defaults to `true`. If set to
     * `false`, this will generate a simple HTML list of links.
     */
    public bool $asDropdown = true;

    /**
     * @var ?string the pjax container identifier inside which this menu is being rendered. If set the jQuery export
     * menu plugin will get auto initialized on pjax request completion.
     */
    public ?string $pjaxContainerId;

    /**
     * @var bool whether to clear all previous / parent buffers. Defaults to `false`.
     */
    public bool $clearBuffers = false;

    /**
     * @var array the HTML attributes for the export button menu. Applicable only if [[asDropdown]] is set to `true`.
     * The following special options are available:
     * - `label`: _string_, defaults to empty string
     * - `icon`: _string_, defaults to `<i class="glyphicon glyphicon-export"></i>` for BS3 and
     * `<i class="fas fa-external-link-alt"></i>` for BS4
     * - `title`: _string_, defaults to `Export data in selected format`.
     * - `menuOptions`: _array_, the HTML attributes for the dropdown menu.
     * - `itemsBefore`: _array_, any additional items that will be merged/prepended before with the export dropdown list.
     *   This should be similar to the `items` property as supported by [[ButtonDropdown]] widget. Note the page export
     *   items will be automatically generated based on settings in the [[exportConfig]] property.
     * - `itemsAfter`: _array_, any additional items that will be merged/appended after with the export dropdown list. This
     *   should be similar to the `items` property as supported by [[ButtonDropdown]] widget. Note the
     *   page export items will be automatically generated based on settings in the `exportConfig` property.
     */
    public array $dropdownOptions = [];

    /**
     * @var bool whether to initialize data provider and clear models before rendering. Defaults to `false`.
     */
    public bool $initProvider = false;

    /**
     * @var bool whether to show a column selector to select columns for export. Defaults to `true`.
     * This is applicable only if [[asDropdown]] is set to `true`. Else this property is ignored.
     */
    public bool $showColumnSelector = true;

    /**
     * @var bool enable or disable cell formatting by auto detecting the grid column alignment and format.
     * If set to `false` the format will not be applied but improve performance configured in [[dynagridOptions]].
     */
    public bool $enableAutoFormat = true;

    /**
     * @var array the configuration of the column names in the column selector. Note: column names will be generated
     * automatically by default. Any setting in this property will override the auto-generated column names. This
     * list should be setup as `$key => $value` where:
     * - `$key`: _integer_, is the zero based index of the column as set in `$columns`.
     * - `$value`: _string_, is the column name/label you wish to display in the column selector.
     */
    public array $columnSelector = [];

    /**
     * @var array the HTML attributes for the column selector dropdown button. The following special options are
     * recognized:
     * - `label`: _string_, defaults to empty string.
     * - `icon`: _string_, defaults to `<i class="glyphicon glyphicon-list"></i>` for BS3 and
     * `<i class="fas fa-list"></i>` for BS4
     * - `title`: _string_, defaults to `Select columns for export`.
     */
    public array $columnSelectorOptions = [];

    /**
     * @var array the HTML attributes for the column selector menu list.
     */
    public array $columnSelectorMenuOptions = [];

    /**
     * @var array the settings for the toggle all checkbox to check / uncheck the columns as a batch. Should be setup as
     * an associative array which can have the following keys:
     * - `show`: _boolean_, whether the batch toggle checkbox is to be shown. Defaults to `true`.
     * - `label`: _string_, the label to be displayed for toggle all. Defaults to `Select Columns`.
     * - `options`: _array_, the HTML attributes for the toggle label text. Defaults to `['class'=>'kv-toggle-all']`
     */
    public array $columnBatchToggleSettings = [];

    /**
     * @var array, HTML attributes for the container to wrap the widget. Defaults to:
     * `['class'=>'btn-group', 'role'=>'group']`
     */
    public array $container = ['class' => 'btn-group', 'role' => 'group'];

    /**
     * @var string, the template for rendering the content in the container. This will be parsed only if `asDropdown`
     * is `true`. The following tokens will be replaced:
     * - `{columns}`: will be replaced with the column selector dropdown
     * - `{menu}`: will be replaced with export menu dropdown
     */
    public string $template = "{columns}\n{menu}";

    /**
     * @var integer timeout for the export function (in seconds), if timeout is < 0, the default PHP timeout will be used.
     */
    public int $timeout = -1;

    /**
     * @var array the HTML attributes for the export form.
     */
    public array $exportFormOptions = [];

    /**
     * @var array the configuration of additional hidden inputs that will be rendered with the export form. This will be
     * set as an array of hidden input `$name => $setting` pairs where:
     * - `$name`: _string_, is the hidden input name attribute.
     * - `$setting`: _array_, containing the following settings:
     *    - `value`: _string_, the value of the hidden input. Defaults to NULL.
     *    - `options`: _array_, the HTML attributes for the hidden input. Defaults to an empty array `[]`.
     *
     * An example of how you can configure this property is shown below:
     *
     * ```
     * [
     *    'inputName1' => ['value' => 'inputValue1'],
     *    'inputName2' => ['value' => 'inputValue2', 'options' => ['id' => 'inputId2']],
     * ]
     * ```
     */
    public array $exportFormHiddenInputs = [];

    /**
     * @var array the selected column indexes for export. If not set this will default to all columns.
     */
    public array $selectedColumns;

    /**
     * @var array the column indexes for export that will be disabled for selection in the column selector.
     */
    public array $disabledColumns = [];

    /**
     * @var array the column indexes for export that will be hidden for selection in the column selector, but will
     * still be displayed in export output.
     */
    public array $hiddenColumns = [];

    /**
     * @var array the column indexes for export that will not be exported at all nor will they be shown in the column
     * selector
     */
    public array $noExportColumns = [];

    /**
     * @var string the view file for rendering the columns selection
     */
    public string $exportColumnsView;

    /**
     * @var bool whether to use font awesome icons for rendering the icons as defined in [[exportConfig]]. If set to
     * `true`, you must load the FontAwesome CSS separately in your application.
     */
    public bool $fontAwesome = false;

    /**
     * @var array the export configuration. The array keys must be the one of the `format` constants (CSV, HTML, TEXT,
     * EXCEL, PDF) and the array value is a configuration array consisting of these settings:
     * - `label`: _string_, the label for the export format menu item displayed
     * - `icon`: _string_, the glyphicon or font-awesome name suffix to be displayed before the export menu item label. If
     *   set to an empty string_, this will not be displayed.
     * - `iconOptions`: _array_, HTML attributes for export menu icon.
     * - `linkOptions`: _array_, HTML attributes for each export item link.
     * - `filename`: _string_, the base file name for the generated file. Defaults to 'grid-export'. This will be used to generate
     *   a default file name for downloading.
     * - `extension`: _string_, the extension for the file name
     * - `alertMsg`: _string_, the message prompt to show before saving. If this is empty or not set it will not be
     *   displayed.
     * - `mime`: _string_, the mime type (for the file format) to be set before downloading.
     * - `writer`: _string_, the PhpSpreadsheet writer type
     * - `options`: _array_, HTML attributes for the export menu item.
     */
    public $exportConfig = [];

    /**
     * @var string the request parameter ($_GET or $_POST) that will be submitted during export. If not set this will
     *  be auto generated. This should be unique for each export menu widget (for multiple export menu widgets on
     *  same page).
     */
    public ?string $exportRequestParam;

    /**
     * @var string the export type input parameter for export form
     */
    public string $exportTypeParam = 'export_type';

    /**
     * @var string the export columns input parameter for export form
     */
    public string $exportColsParam = 'export_columns';

    /**
     * @var string the column selector flag parameter for export form
     */
    public string $colSelFlagParam = 'column_selector_enabled';

    /**
     * @var array the output style configuration options for each data cell. It must be the style configuration
     * array as required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public array $styleOptions = [];

    /**
     * @var array the output style configuration options for the header row. It must be the style configuration array as
     * required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public array $headerStyleOptions = [];

    /**
     * @var array the output style configuration options for the entire spreadsheet box range. It must be the style
     * configuration array as required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public array $boxStyleOptions = [];

    /**
     * @var array an array of rows to prepend in front of the grid used to create things like a title. Each array
     * should be set with the following settings:
     * - value: string, the value of the merged row
     * - cellFormat: string|null, the explicit cell format to apply.
     * - styleOptions: array, array of configuration options to set the style. See $styleOptions on how to configure.
     */
    public array $contentBefore = [];

    /**
     * @var array an array of rows to append after the footer row. Each array
     * should be set with the following settings:
     * - value: string, the value of the merged row
     * - cellFormat: string|null, the explicit cell format to apply
     * - styleOptions: array, array of configuration options to set the style. See $styleOptions on how to configure.
     */
    public array $contentAfter = [];

    /**
     * @var bool whether to auto-size the excel output column widths. Defaults to `true`.
     */
    public bool $autoWidth = true;

    /**
     * @var array The number of characters for each column - used to approximate an autoWidth for openspout
     */
    protected array $autoWidthColumns = [];

    /**
     * @var string encoding for the downloaded file header. Defaults to [[ENCODING_UTF8]].
     */
    public string $encoding = self::ENCODING_UTF8;

    /**
     * @var string the exported output file name. Defaults to 'grid-export';
     */
    public string $filename;

    /**
     * @var string the folder to save the exported file. Defaults to '@app/runtime/export/'. If the specified folder
     * does not exist, the extension will attempt to create it - else an exception will be thrown.
     */
    public string $folder = '@app/runtime/export';

    /**
     * @var string the web accessible path for the saved file location. This property will be parsed only if [[stream]]
     * is false. Note the [[afterSaveView]] property that will render the displayed file link.
     */
    public string $linkPath = '/runtime/export';

    /**
     * @var string the name of the file to be appended to [[linkPath]] to generate the complete link. If not set, this
     * will default to the [[filename]].
     */
    public string $linkFileName;

    /**
     * @var bool whether to stream output to the browser.
     */
    public bool $stream = true;

    /**
     * @var boolean whether to delete file after saving file to [[folder]] and when [[stream]] is `false`. This property
     * will be validated only when [[stream]] is `false`.
     */
    public bool $deleteAfterSave = false;

    /**
     * @var string|false the view file to show details of exported file link. This property will be validated only when
     * [[stream]] is `false`. You can set this to `false` to not display any file link details for view. This defaults
     * to the `_view` PHP file in the `views` folder of the extension.
     */
    public string|false $afterSaveView;

    /**
     * @var integer  fetch models from the dataprovider using batches of this size. Set this to `0` (the default) to
     * disable. If `$dataProvider` does not have a pagination object, this parameter is ignored. Setting this
     * property helps reduce memory overflow issues by allowing parsing of models in batches, rather than fetching
     * all models in one go.
     */
    public int $batchSize = 0;

    /**
     * @var array, the configuration of various messages that will be displayed at runtime:
     * - allowPopups: string, the message to be shown to disable browser popups for download. Defaults to `Disable any
     *   popup blockers in your browser to ensure proper download.`.
     * - confirmDownload: string, the message to be shown for confirming to proceed with the download. Defaults to `Ok
     *   to proceed?`.
     * - downloadProgress: string, the message to be shown in a popup dialog when download request is executed.
     *   Defaults to `Generating file. Please wait...`.
     * - downloadComplete: string, the message to be shown in a popup dialog when download request is completed.
     *   Defaults to `All done! Click anywhere here to close this window, once you have downloaded the file.`.
     */
    public array $messages = [];

    /**
     * @var array the document's properties. Only works when using openspout >= 4.28.0.
     */
    public array $docProperties = [];

    /**
     * @var boolean enable dynagrid for column selection. If set to `true` the inbuilt export menu column selector
     * functionality will be disabled and not rendered and column settings for dynagrid will be used as per settings
     * configured in [[dynagridOptions]].
     */
    public bool $dynagrid = false;

    /**
     * @var array dynagrid widget options. Applicable only if [[dynagrid]] is set to `true`.
     */
    public array $dynagridOptions = ['options' => ['id' => 'dynagrid-export-menu']];

    /**
     * @var array the PhpSpreadsheet style configuration for a grouped grid row
     */
    public array $groupedRowStyle = [
        'font' => [
            'bold'  => false,
            'color' => [
                'argb' => 'FF000080',
            ],
        ],
        'fill' => [
            'type'  => 'solid',
            'color' => [
                'argb' => 'FFFFFFFF',
            ],
        ],
    ];

    /**
     * @var array|null new supplement sheets to be created. Required for data validation in excel. An example setting:
     * ```
     * 'supplementSheets' => ['city' => $citiesKeyValueArray, 'state' => $stateKeyValueArray],
     * ```
     * If set to empty or null will be ignored.
     */
    public array|null $supplementSheets = null;

    /**
     * @var string the data output format type. Defaults to `ExportMenu::FORMAT_EXCEL_X`.
     */
    public string $exportType = self::FORMAT_EXCEL_X;

    /**
     * @var bool flag to identify if download is triggered
     * - Auto filtering of columns is lost.
     */
    public bool $triggerDownload = false;

    /**
     * @var BaseDataProvider the modified data provider for usage with export.
     */
    protected BaseDataProvider $_provider;

    /**
     * @var array the default export configuration
     */

    protected array $_defaultExportConfig = [];

    /**
     * @var Writer|OpenspoutCsvWriter|OpenspoutOdsWriter object instance
     */
    protected Writer|OpenspoutCsvWriter|OpenspoutOdsWriter $_objOpenspoutWriter;

    /**
     * @var Options|OpenspoutCsvOptions|OpenspoutOdsOptions object instance
     */
    protected Options|OpenspoutCsvOptions|OpenspoutOdsOptions $_objOpenspoutOptions;

    /**
     * @var Sheet|null object instance
     */
    protected Sheet|null $_objOpenspoutSheet = null;

    /**
     * @var int the header beginning row
     */
    protected int $_headerBeginRow = 1;

    /**
     * @var int the table beginning row
     */
    protected int $_beginRow = 1;

    /**
     * @var int the current table end row
     */
    protected int $_endRow = 0;

    /**
     * @var int the current table end column
     */
    protected int $_endCol = 1;

    /**
     * @var bool whether the column selector is enabled
     */
    protected bool $_columnSelectorEnabled;

    /**
     * @var array the visble columns for export
     */
    protected array $_visibleColumns;

    /**
     * @var array columns to be grouped
     */
    protected array $_groupedColumn = [];

    /**
     * @var array|null grouped row values
     */
    protected array|null $_groupedRow = null;

    /**
     * @var string the data output format type. Defaults to `ExportMenu::FORMAT_EXCEL_X`.
     */
    protected string $_exportType;

    /**
     * @var bool private flag that will use $_POST [[exportRequestParam]] setting if available or use the
     * [[triggerDownload]] setting
     */
    protected bool $_triggerDownload;

    /**
     * Appends slash to path if it does not exist
     *
     * @param string $path
     * @param string $s the path separator
     *
     * @return string
     */
    protected static function slash($path, $s = DIRECTORY_SEPARATOR)
    {
        $path = trim($path);
        if (substr($path, -1) !== $s) {
            $path .= $s;
        }

        return $path;
    }

    /**
     * Determine the target directory for this export.
     * @param $config
     * @return string
     */
    protected function getTargetDirectory($config)
    {
        $this->folder = trim(Yii::getAlias($this->folder));
        if (!file_exists($this->folder) && !mkdir($this->folder, 0777, true)) {
            throw new InvalidConfigException(
                "Invalid permissions to write to '{$this->folder}' as set in `ExportMenu::folder` property."
            );
        }
        $filename = static::sanitize($this->filename);
        return self::slash($this->folder) . $filename . '.' . $config['extension'];
    }

    /**
     * Returns an excel column name.
     *
     * @param integer $index the column index number
     *
     * @return string
     */
    protected static function columnName($index)
    {
        $i = (int)($index) - 1;
        if ($i >= 0 && $i < 26) {
            return chr(ord('A') + $i);
        }
        if ($i > 25) {
            return (self::columnName($i / 26)) . (self::columnName($i % 26 + 1));
        }

        return 'A';
    }

    /**
     * @inheritdoc
     */
    public function init()
    {
        $this->initSettings();
        parent::init();
    }

    /**
     * @inheritdoc
     */
    public function run()
    {
        $this->initI18N(__DIR__);
        $this->initColumnSelector();
        $this->initNoExportColumns();
        $this->setVisibleColumns();
        $this->initExport();
        if (!$this->_triggerDownload) {
            $this->registerAssets();
            echo $this->renderExportMenu();

            return;
        }
        if ($this->timeout >= 0) {
            set_time_limit($this->timeout);
        }
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        $file = $this->getTargetDirectory($config);
        $this->initOpenspout();
        if ($this->stream) {
            $this->clearOutputBuffers();
            $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
            $extension = ArrayHelper::getValue($config, 'extension', 'xlsx');
            header('Expires: Sat, 26 Jul 1997 05:00:00 GMT');
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
            $this->_objOpenspoutWriter->openToBrowser($this->filename . '.' . $extension);
        } else {
            $this->_objOpenspoutWriter->openToFile($file);
        }
        $this->initOpenspoutSheetView();
        $this->generateBeforeContent();
        $this->generateHeader();
        $this->generateBody();
        if (!empty($this->supplementSheets)) {
            $this->createSupplementSheets();
        }
        $this->_objOpenspoutWriter->setCurrentSheet($this->_objOpenspoutSheet);
        $row = $this->generateFooter();
        $this->generateAfterContent($row);
        if ($this->_objOpenspoutOptions instanceof AbstractOptions) {
            $columns = $this->getVisibleColumns();
            $col = 0;
            // ODS and XLSX files scale widths differently, so use different scaling factor
            $factor = $this->_exportType === self::FORMAT_ODS ? 5.0 : 1.0;
            foreach ($columns as $column) {
                $col++;
                if (isset($column->options['width'])) {
                    $this->_objOpenspoutOptions->setColumnWidth($column->options['width'] * $factor, $col);
                    continue;
                }
                if ($this->autoWidth) {
                    $this->_objOpenspoutOptions->setColumnWidth($this->autoWidthColumns[$col] * 1.2 * $factor, $col);
                }
            }
        }
        $this->_objOpenspoutWriter->close();
        if ($this->stream) {
            exit();
        }

        $this->registerAssets();
        echo $this->renderExportMenu();
        if ($this->_triggerDownload && $this->afterSaveView !== false) {
            $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
            if (!empty($config)) {
                $l = $this->linkFileName;
                $fileName = (!isset($l) || $l === '' ? $this->filename : $l) . '.' . $config['extension'];
                echo $this->render(
                    $this->afterSaveView,
                    [
                        'notBs3' => !$this->isBs(3),
                        'file'   => $fileName,
                        'icon'   => $config['icon'],
                        'href'   => Url::to([self::slash($this->linkPath, '/') . $fileName]),
                    ]
                );
            }
        }
        $this->cleanup($file);
    }

    /**
     * Initialize export menu settings
     */
    protected function initSettings()
    {
        $this->showPageSummary = false; // force disable page-summary for `ExportMenu` (issue #162)
        $this->_msgCat = 'kvexport';
        if (empty($this->options['id'])) {
            $this->options['id'] = $this->getId();
        }
        if (empty($this->exportRequestParam)) {
            $this->exportRequestParam = 'exportFull_' . $this->options['id'];
        }
        $path = '@vendor/kartik-v/yii2-export/src/views';
        if (!isset($this->exportColumnsView)) {
            $this->exportColumnsView = "{$path}/_columns";
        }
        if (!isset($this->afterSaveView)) {
            $this->afterSaveView = "{$path}/_view";
        }
        $this->_columnSelectorEnabled = $this->showColumnSelector && $this->asDropdown;
        $request = Yii::$app->request;
        if ($request instanceof Request) {
            $this->_triggerDownload = $request->post($this->exportRequestParam, $this->triggerDownload);
            $this->_exportType = $request->post($this->exportTypeParam, $this->exportType);
        } else {
            $this->_triggerDownload = $this->triggerDownload;
            $this->_exportType = $this->exportType;
        }
        if (!$this->stream) {
            $this->target = self::TARGET_SELF;
        }
        if ($this->_triggerDownload) {
            if ($this->stream) {
                Yii::$app->controller->layout = false;
            }
            $this->_columnSelectorEnabled = $request instanceof Request ?
                $request->post($this->colSelFlagParam, $this->_columnSelectorEnabled) :
                $this->_columnSelectorEnabled;
            $this->initSelectedColumns();
        }
        if ($this->dynagrid) {
            $this->_columnSelectorEnabled = false;
            $options = $this->dynagridOptions;
            $options['columns'] = $this->columns;
            if (!isset($options['storage'])) {
                $options['storage'] = DynaGrid::TYPE_DB;
            }
            $options['gridOptions']['dataProvider'] = $this->dataProvider;
            $dynagrid = new DynaGrid($options);
            $this->columns = $dynagrid->getColumns();
        }
    }

    /**
     * Create supplement sheets, usually from dropDownList used by gridView. Needed for using dataValidation.
     *
     * @throws IOException
     * @throws WriterNotOpenedException
     * @throws InvalidSheetNameException
     */
    protected function createSupplementSheets()
    {
        if (!in_array($this->_exportType, [self::FORMAT_EXCEL_X, self::FORMAT_ODS], true)) {
            return;
        }
        foreach ($this->supplementSheets as $sheetName => $sheetData) {
            $newSheet = $this->_objOpenspoutWriter->addNewSheetAndMakeItCurrent();
            $newSheet->setName($sheetName);
            $this->_objOpenspoutWriter->addRow(Row::fromValues([OpenspoutCell::fromValue(Yii::t('kvexport', 'Key')), OpenspoutCell::fromValue(Yii::t('kvexport', 'Value'))]));
            foreach ($sheetData as $key => $value) {
                $this->_objOpenspoutWriter->addRow(Row::fromValues([$key, $value]));
            }
        }
    }

    /**
     * Initializes export settings
     */
    protected function initExport()
    {
        $this->_provider = clone($this->dataProvider);
        if ($this->batchSize && $this->_provider->pagination) {
            $this->_provider->pagination = clone($this->dataProvider->pagination);
            $this->_provider->pagination->pageSize = $this->batchSize;
            $this->_provider->refresh();
            if (Yii::$app->request->getBodyParam('exportFull_export')) {
                $this->_provider->pagination->page = null;
                Yii::$app->request->setQueryParams([$this->_provider->pagination->pageParam => 1]);
            }
        } else {
            $this->_provider->pagination = false;
        }
        if ($this->initProvider) {
            $this->_provider->prepare(true);
        }
        $this->setDefaultStyles('header');
        $this->setDefaultStyles('box');
        $this->filterModel = null;
        $this->setDefaultExportConfig();
        $this->exportConfig = ArrayHelper::merge($this->_defaultExportConfig, $this->exportConfig);
        if (!isset($this->filename)) {
            $this->filename = Yii::t('kvexport', 'grid-export');
        }
        Html::addCssClass($this->exportFormOptions, 'kv-export-full-form');
        if (!isset($this->exportFormOptions['id'])) {
            $this->exportFormOptions['id'] = $this->options['id'] . '-export-form';
        }
        $this->_provider->refresh();
    }

    /**
     * Renders the export menu widget.
     *
     * @return string the export menu markup
     * @throws InvalidConfigException
     * @throws Exception
     */
    protected function renderExportMenu()
    {
        $items = $this->asDropdown ? [] : '';
        foreach ($this->exportConfig as $format => $settings) {
            if (!isset($settings) || $settings === false) {
                continue;
            }
            $label = '';
            if (isset($settings['icon'])) {
                $iconOptions = ArrayHelper::getValue($settings, 'iconOptions', []);
                Html::addCssClass($iconOptions, $settings['icon']);
                $label = Html::tag('i', '', $iconOptions) . ' ';
            }
            if (isset($settings['label'])) {
                $label .= $settings['label'];
            }
            $fmt = strtolower($format);
            $linkOptions = ArrayHelper::getValue($settings, 'linkOptions', []);
            $linkOptions['id'] = $this->options['id'] . '-' . $fmt;
            $linkOptions['data-format'] = $format;
            $options = ArrayHelper::getValue($settings, 'options', []);
            Html::addCssClass($linkOptions, "export-full-{$fmt}");
            if ($this->asDropdown) {
                $items[] = [
                    'label'       => $label,
                    'url'         => '#',
                    'linkOptions' => $linkOptions,
                    'options'     => $options,
                ];
            } else {
                $tag = ArrayHelper::remove($options, 'tag', 'li');
                if ($tag !== false) {
                    $items .= Html::tag($tag, Html::a($label, '#', $linkOptions), $options) . "\n";
                } else {
                    $items .= Html::a($label, '#', $linkOptions) . "\n";
                }
            }
        }
        if ($this->asDropdown) {
            $this->replacePart('template', '{menu}', [$this, 'renderDropdownMenu'], [$items]);
            $this->replacePart('template', '{columns}', [$this, 'renderColumnSelector']);

            return Html::tag('div', $this->template, $this->container);
        }

        return $items;
    }

    /**
     * Renders the dropdown menu button and items.
     *
     * @param array $items
     * @return string
     * @throws InvalidConfigException|Throwable
     */
    protected function renderDropdownMenu($items)
    {
        Html::addCssClass($this->dropdownOptions, ['btn', $this->getDefaultBtnCss()]);
        $notBs3 = !$this->isBs(3);
        $iconCss = $notBs3 ? 'fas fa-external-link-alt' : 'glyphicon glyphicon-export';
        $icon = ArrayHelper::remove($this->dropdownOptions, 'icon', '<i class="' . $iconCss . '"></i>');
        $label = ArrayHelper::remove($this->dropdownOptions, 'label');
        $label = $label === null ? $icon : $icon . ' ' . $label;
        if (!isset($this->dropdownOptions['title'])) {
            $this->dropdownOptions['title'] = Yii::t('kvexport', 'Export data in selected format');
        }
        $menuOptions = ArrayHelper::remove($this->dropdownOptions, 'menuOptions', []);
        $itemsBefore = ArrayHelper::remove($this->dropdownOptions, 'itemsBefore', []);
        $itemsAfter = ArrayHelper::remove($this->dropdownOptions, 'itemsAfter', []);
        $items = ArrayHelper::merge($itemsBefore, $items, $itemsAfter);
        $opts = [
            'label'       => $label,
            'dropdown'    => ['items' => $items, 'encodeLabels' => false, 'options' => $menuOptions,],
            'encodeLabel' => false,
        ];

        if (!isset($this->exportContainer['class'])) {
            $this->exportContainer['class'] = 'btn-group';
        }
        /**
         * @var Widget $class
         */
        $class = $this->getDropdownClass(true);
        if ($notBs3) {
            $opts['buttonOptions'] = $this->dropdownOptions;
            $opts['renderContainer'] = false;
            $out = Html::tag('div', $class::widget($opts), $this->exportContainer);
        } else {
            $opts['options'] = $this->dropdownOptions;
            $opts['containerOptions'] = $this->exportContainer;
            $out = $class::widget($opts);
        }

        return $out;
    }

    /**
     * Renders the columns selector
     *
     * @return string the column selector markup
     * @throws Exception
     */
    protected function renderColumnSelector()
    {
        if (!$this->_columnSelectorEnabled) {
            return '';
        }

        return $this->render(
            $this->exportColumnsView,
            [
                'id'              => $this->options['id'],
                'notBs3'          => !$this->isBs(3),
                'isBs4'           => $this->isBs(4),
                'options'         => $this->columnSelectorOptions,
                'menuOptions'     => $this->columnSelectorMenuOptions,
                'columnSelector'  => $this->columnSelector,
                'batchToggle'     => $this->columnBatchToggleSettings,
                'selectedColumns' => $this->selectedColumns,
                'disabledColumns' => $this->disabledColumns,
                'hiddenColumns'   => $this->hiddenColumns,
                'noExportColumns' => $this->noExportColumns,
            ]
        );
    }

    /**
     * Initializes Openspout Object Instance
     */
    protected function initOpenspout()
    {
        if ($this->_exportType === self::FORMAT_EXCEL_X) {
            $this->_objOpenspoutOptions = new Options();
            // Properties is only supported from openspout 4.28.0, but we need to support down to 4.26.0 for php8.1 support
            if (class_exists('Properties')) {
                $creator = $title = $subject = $category = $keywords = '';
                $description = Yii::t('kvexport', 'Grid export generated by Krajee ExportMenu widget (yii2-export)');
                $lastModifiedBy = 'krajee';
                extract($this->docProperties);
                $properties = new Properties(
                    title         : $title,
                    subject       : $subject,
                    creator       : $creator,
                    lastModifiedBy: $lastModifiedBy,
                    keywords      : $keywords,
                    description   : $description,
                    category      : $category
                );
                $this->_objOpenspoutOptions->setProperties($properties);
            }
            $this->_objOpenspoutWriter = new Writer($this->_objOpenspoutOptions);
        } elseif ($this->_exportType === self::FORMAT_ODS) {
            $this->_objOpenspoutOptions = new OpenspoutOdsOptions();
            $this->_objOpenspoutWriter = new OpenspoutOdsWriter($this->_objOpenspoutOptions);
        } else {
            $this->_objOpenspoutOptions = new OpenspoutCsvOptions();
            $this->_objOpenspoutOptions->FIELD_DELIMITER = $this->getSetting('delimiter', "\t");
            $this->_objOpenspoutWriter = new OpenspoutCsvWriter($this->_objOpenspoutOptions);
        }
    }

    /**
     * Initialise the Openspout sheet view.
     * @return void
     * @throws InvalidArgumentException
     * @throws WriterNotOpenedException
     */
    protected function initOpenspoutSheetView()
    {
        if (!in_array($this->_exportType, [self::FORMAT_EXCEL_X, self::FORMAT_ODS], true)) {
            return;
        }
        $sheetView = new SheetView();
        // freeze header row
        $sheetView->setFreezeRow(2 + count($this->contentBefore));
        $this->_objOpenspoutSheet = $this->_objOpenspoutWriter->getCurrentSheet();
        $this->_objOpenspoutSheet->setSheetView($sheetView);
    }

    /**
     * Generates the before content at the top of the exported sheet
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     * @throws WriterNotOpenedException
     */
    protected function generateBeforeContent()
    {
        foreach ($this->contentBefore as $contentBefore) {
            $format = ArrayHelper::getValue($contentBefore, 'cellFormat');
            $opts = $this->getStyleOpts($contentBefore);
            $style = OpenspoutHelper::createStyleFromPhpSpreadsheetOptions($opts);
            $style->setFormat($format);
            $cell = OpenspoutCell::fromValue($contentBefore['value'], $style);
            $this->_objOpenspoutWriter->addRow(new Row([$cell]));
            $this->_beginRow++;
        }
    }

    /**
     * Generates the output data header content.
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws WriterNotOpenedException
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     */
    protected function generateHeader()
    {
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return;
        }
        $styleOpts = ArrayHelper::getValue($this->headerStyleOptions, $this->_exportType, []);

        $this->_endCol = 0;
        $openspoutCells = [];
        foreach ($columns as $column) {
            $opts = $styleOpts;
            $this->_endCol++;
            /**
             * @var DataColumn $column
             */
            $head = ($column instanceof DataColumn) ? $this->getColumnHeader($column) : $column->header;
            $format = ArrayHelper::remove($column->headerOptions, 'cellFormat');
            if (isset($column->hAlign) && !isset($opts['alignment']['horizontal'])) {
                $opts['alignment']['horizontal'] = $column->hAlign;
            }
            if (isset($column->vAlign) && !isset($opts['alignment']['vertical'])) {
                $opts['alignment']['vertical'] = $column->vAlign;
            }
            $style = OpenspoutHelper::createStyleFromPhpSpreadsheetOptions(array_replace_recursive($this->getBoxStyleArrayForCell($this->_endCol, true, false), $opts));
            if (!empty($format)) {
                $style->setFormat($format);
            }
            $this->autoWidthColumns[$this->_endCol] = strlen((string)$head);
            $openspoutCells[] = OpenspoutCell::fromValue($head, $style);
        }
        $this->_objOpenspoutWriter->addRow(new Row($openspoutCells));
        for ($i = $this->_headerBeginRow; $i < ($this->_beginRow); $i++) {
            $this->mergeCells(1, $i, $this->_endCol, $i);
        }
    }

    /**
     * Gets the visible columns for export
     *
     * @return array the columns configuration
     */
    protected function getVisibleColumns()
    {
        if (!isset($this->_visibleColumns)) {
            $this->setVisibleColumns();
        }

        return $this->_visibleColumns;
    }

    /**
     * Add any action columns to the noExportColumns field.
     */
    protected function initNoExportColumns()
    {
        foreach ($this->columns as $key => $column) {
            $isActionColumn = $column instanceof ActionColumn;
            $isNoExport = in_array($key, $this->noExportColumns, false) ||
                ($this->showColumnSelector && !in_array($key, $this->selectedColumns, false));
            if ($isActionColumn && !$isNoExport) {
                $this->noExportColumns[] = $key;
            }
        }
    }

    /**
     * Sets visible columns for export
     */
    protected function setVisibleColumns()
    {
        $columns = [];
        foreach ($this->selectedColumns as $key) {
            $column = $this->columns[$key];
            if (
                !empty($column->hiddenFromExport)
                || $column instanceof ActionColumn
                || in_array($key, $this->noExportColumns, false)
            ) {
                continue;
            }
            $columns[] = $column;
        }
        $this->_visibleColumns = $columns;
    }

    /**
     * Gets the column header content
     *
     * @param DataColumn $col
     *
     * @return string
     */
    protected function getColumnHeader($col)
    {
        if ($col->header !== null || ($col->label === null && $col->attribute === null)) {
            return trim($col->header) !== '' ? $col->header : $col->grid->emptyCell;
        }
        $provider = $this->dataProvider;
        if ($col->label === null) {
            if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
                $model = new $provider->query->modelClass;
                $label = $model->getAttributeLabel($col->attribute);
            } else {
                $models = $provider->getModels();
                if (($model = reset($models)) instanceof Model) {
                    $label = $model->getAttributeLabel($col->attribute);
                } else {
                    $label = Inflector::camel2words($col->attribute);
                }
            }
        } else {
            $label = $col->label;
        }

        return $label;
    }

    /**
     * Generates the output data body content.
     *
     * @return integer the number of output rows.
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     * @throws WriterNotOpenedException
     */
    protected function generateBody()
    {
        $this->_endRow = 0;
        $columns = $this->getVisibleColumns();
        $models = array_values($this->_provider->getModels());
        if (count($columns) === 0) {
            $this->_objOpenspoutWriter->addRow(Row::fromValues([$this->emptyText]));
            return 0;
        }
        // do not execute multiple COUNT(*) queries
        $totalCount = $this->_provider->getTotalCount();
        // we need to keep track of the current row to know when we've arrived at the last row
        $currentRow = 1;
        $this->findGroupedColumn();
        while (count($models) > 0) {
            $keys = $this->_provider->getKeys();
            foreach ($models as $index => $model) {
                $key = $keys[$index];
                $isLastRow = $currentRow === $totalCount;
                if ($isLastRow) {
                    //a little hack to generate last grouped footer
                    $this->checkGroupedRow($model, $models[0], $key, $this->_endRow + 1);
                } elseif (isset($models[$index + 1])) {
                    $this->checkGroupedRow($model, $models[$index + 1], $key, $this->_endRow + 1);
                }
                $this->generateRow($model, $key, $this->_endRow, $isLastRow && $this->_groupedRow === null);
                $this->_endRow++;
                if (!is_null($this->_groupedRow)) {
                    $this->_endRow++;
                    $cells = array_map(
                        function ($value, $idx) use ($isLastRow) {
                            $groupedRowStyle = OpenspoutHelper::createStyleFromPhpSpreadsheetOptions(
                                array_replace_recursive(
                                    $this->getBoxStyleArrayForCell($idx + 1, false, $isLastRow),
                                    $this->groupedRowStyle
                                )
                            );
                            return OpenspoutCell::fromValue($value, $groupedRowStyle);
                        },
                        $this->_groupedRow,
                        array_keys($this->_groupedRow)
                    );
                    $this->_objOpenspoutWriter->addRow(new Row($cells));
                    $this->_groupedRow = null;
                }
                $currentRow++;
            }
            if ($this->_provider->pagination) {
                $this->_provider->pagination->page++;
                $this->_provider->refresh();
                $this->_provider->setTotalCount($totalCount);
                $models = array_values($this->_provider->getModels());
            } else {
                $models = [];
            }
        }
        $this->generateBox();

        return $this->_endRow;
    }

    /**
     * Generates an output data row with the given data model and key.
     *
     * @param mixed $model the data model to be rendered
     * @param mixed $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     * @param bool $isLastRow Whether this is the last row in the table.
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     * @throws WriterNotOpenedException
     */
    protected function generateRow($model, $key, $index, $isLastRow = false)
    {
        /**
         * @var Column $column
         */
        $this->_endCol = 0;
        $openspoutCells = [];
        foreach ($this->getVisibleColumns() as $column) {
            $value = null;
            if ($column instanceof SerialColumn) {
                $value = $index + 1;
                $pagination = $column->grid->dataProvider->getPagination();
                if ($pagination !== false) {
                    $value += $pagination->getOffset();
                }
            } elseif (isset($column->content)) {
                $value = call_user_func($column->content, $model, $key, $index, $column);
            } elseif (method_exists($column, 'getDataCellValue')) {
                $value = $column->getDataCellValue($model, $key, $index);
            } elseif (isset($column->attribute)) {
                $value = ArrayHelper::getValue($model, $column->attribute, '');
            }
            $this->_endCol++;
            $yiiFormat = $this->enableFormatter && isset($column->format) ? $column->format : 'raw';
            if (isset($value) && $value !== '' && isset($yiiFormat)) {
                $value = $this->formatter->format($value, $yiiFormat);
            } else {
                $value = '';
            }
            $contentOptions = $column->contentOptions;
            if (is_callable($contentOptions)) {
                $contentOptions = $contentOptions($model, $key, $index, $column);
            }

            //20201026 Scott: To avoid 'Closure object cannot have properties' error
            try {
                $format = ArrayHelper::getValue($contentOptions, 'cellFormat');
            } catch (Exception|Throwable $e) {
                $format = null;
            }

            $opts = $this->getBoxStyleArrayForCell($this->_endCol, false, $isLastRow);
            if ($format === null && $this->enableAutoFormat) {
                $opts = array_replace_recursive($opts, $this->getAutoFormattedOpts($model, $key, $index, $column));
            }
            $style = OpenspoutHelper::createStyleFromPhpSpreadsheetOptions($opts);
            if ($format !== null) {
                $style->setFormat($format);
            }
            $length = match (true) {
                is_string($value)                    => strlen($value),
                $value instanceof \DateTimeInterface => 10,
                default                              => 0,
            };
            $this->autoWidthColumns[$this->_endCol] = max($length, $this->autoWidthColumns[$this->_endCol]);
            $openspoutCells[] = OpenspoutCell::fromValue($value, $style);
        }
        $this->_objOpenspoutWriter->addRow(new Row($openspoutCells));
    }

    /**
     * Generates the output footer row after a specific row number
     *
     * @return integer the row number after which the footer is to be generated
     * @throws IOException
     * @throws WriterNotOpenedException
     */
    protected function generateFooter()
    {
        $row = $this->_endRow + $this->_beginRow;
        $footerExists = false;
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return 0;
        }
        $this->_endCol = 0;
        $openspoutCells = [];
        foreach ($this->getVisibleColumns() as $column) {
            $this->_endCol++;
            if ($column->footer) {
                $footerExists = true;
                $footer = trim($column->footer) !== '' ? $column->footer : $column->grid->blankDisplay;
                $format = ArrayHelper::remove($column->footerOptions, 'cellFormat');
                $style = new Style();
                $style->setFormat($format);
                $this->autoWidthColumns[$this->_endCol] = max(strlen($footer), $this->autoWidthColumns[$this->_endCol]);
                $openspoutCells[] = OpenspoutCell::fromValue($footer, $style);
            } else {
                $openspoutCells[] = OpenspoutCell::fromValue('');
            }
        }
        if ($footerExists) {
            $row++;
            $this->_objOpenspoutWriter->addRow(new Row($openspoutCells));
        }

        return $row;
    }

    /**
     * Generates the after content at the bottom of the exported sheet
     *
     * @param integer $row the row number after which the content is to be generated
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     * @throws WriterNotOpenedException
     */
    protected function generateAfterContent(int $row)
    {
        $row++;
        $afterContentBeginRow = $row;
        foreach ($this->contentAfter as $contentAfter) {
            $format = ArrayHelper::getValue($contentAfter, 'cellFormat');
            $opts = $this->getStyleOpts($contentAfter);
            $style = OpenspoutHelper::createStyleFromPhpSpreadsheetOptions($opts);
            $style->setFormat($format);
            $cell = OpenspoutCell::fromValue($contentAfter['value'], $style);
            $this->_objOpenspoutWriter->addRow(new Row([$cell]));
            $row++;
        }
        for ($i = $afterContentBeginRow; $i < $row; $i++) {
            $this->mergeCells(1, $i, $this->_endCol, $i);
        }
    }

    /**
     * Sets default styles
     *
     * @param string $section the php spreadsheet section
     */
    protected function setDefaultStyles($section)
    {
        $defaultStyle = [];
        $opts = '';
        if ($section === 'header') {
            $opts = 'headerStyleOptions';
            $defaultStyle = [
                'font'    => ['bold' => true],
                'fill'    => [
                    'fillType' => 'solid',
                    'color'    => [
                        'argb' => 'FFE5E5E5',
                    ],
                ],
                'borders' => [
                    'outline' => [
                        'borderStyle' => 'medium',
                        'color'       => ['argb' => 'FF000000'],
                    ],
                    'inside'  => [
                        'borderStyle' => 'thin',
                        'color'       => ['argb' => 'FF000000'],
                    ],
                ],
            ];
        } elseif ($section === 'box') {
            $opts = 'boxStyleOptions';
            $defaultStyle = [
                'borders' => [
                    'outline' => [
                        'borderStyle' => 'medium',
                        'color'       => ['argb' => 'FF000000'],
                    ],
                    'inside'  => [
                        'borderStyle' => 'dotted',
                        'color'       => ['argb' => 'FF000000'],
                    ],
                ],
            ];
        }
        if (empty($opts)) {
            return;
        }
        $defaultStyleOptions = [
            self::FORMAT_EXCEL_X => $defaultStyle,
            self::FORMAT_ODS     => $defaultStyle,
        ];
        $this->$opts = array_replace_recursive($defaultStyleOptions, $this->$opts);
    }

    /**
     * Generates the box.
     *
     * Box styles for openspout are created on the fly, so for openspout, this just sets the AutoFilter.
     */
    protected function generateBox()
    {
        $autoFilter = new AutoFilter(0, $this->_beginRow, $this->_endCol - 1, $this->_endRow + $this->_beginRow);
        $this->_objOpenspoutSheet->setAutoFilter($autoFilter);
    }

    /**
     * @param int $col
     * @param bool $isHeader
     * @param bool $isLastRow
     * @return array the style array.
     */
    protected function getBoxStyleArrayForCell(int $col, bool $isHeader, bool $isLastRow): array
    {
        if (!isset($this->boxStyleOptions[$this->_exportType]) && (!$isHeader || !isset($this->headerStyleOptions[$this->_exportType]))) {
            return [];
        }
        $opts = $this->boxStyleOptions[$this->_exportType] ?? [];
        OpenspoutHelper::setInsideAndOutlineBorders($opts, $isHeader, $isLastRow, $col === 1, $col === count($this->getVisibleColumns()));
        if ($isHeader && isset($this->headerStyleOptions[$this->_exportType])) {
            unset($opts['borders']['inside'], $opts['borders']['outline']);
            $opts = array_replace_recursive($opts, $this->headerStyleOptions[$this->_exportType]);
            OpenspoutHelper::setInsideAndOutlineBorders($opts, true, true, $col === 1, $col === count($this->getVisibleColumns()));
        }
        return $opts;
    }

    /**
     * Autoformats a cell by auto detecting the grid column alignment and format
     *
     * @param mixed $model the data model to be rendered
     * @param mixed $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     * @param Column $column
     * @return array|mixed
     */
    protected function getAutoFormattedOpts($model, $key, $index, $column)
    {
        $opts = ArrayHelper::getValue($this->styleOptions, $this->_exportType, []);
        if (isset($column->exportMenuStyle)) {
            $opts = $column->exportMenuStyle;
            if ($opts instanceof Closure) {
                $opts = call_user_func($opts, $model, $key, $index, $column);
            }
        }
        if (isset($column->hAlign) && !isset($opts['alignment']['horizontal'])) {
            $opts['alignment']['horizontal'] = $column->hAlign;
        }
        if (isset($column->vAlign) && !isset($opts['alignment']['vertical'])) {
            $opts['alignment']['vertical'] = $column->vAlign;
        }
        if (isset($column->format) && !isset($opts['numberFormat']) && !($column->format instanceof Closure)) {
            $fmt = (array)$column->format;
            $f = $fmt[0];
            $code = null;
            if ($f === 'integer') {
                $code = '0';
            } elseif ($f === 'percent' || $f === 'decimal' || $f === 'currency') {
                $code = '';
                if ($f === 'currency') {
                    $code = ArrayHelper::getValue($fmt, 1, $this->formatter->currencyCode) . ' ';
                }
                $decimals = ArrayHelper::getValue($fmt, 1, ($f === 'percent' ? 0 : 2));
                $d = (int)$decimals;
                $code .= '#' . $this->formatter->thousandSeparator . '##0';
                if ($d > 0) {
                    $code .= $this->formatter->decimalSeparator . str_repeat('0', $d);
                }
                if ($f === 'percent') {
                    $code .= '%';
                }
            }
            if ($code !== null) {
                $opts['numberFormat'] = ['formatCode' => $code];
            }
        }
        return $opts;
    }

    /**
     * Gets the setting property value for the current export format
     *
     * @param string $key the setting property key for the current export format
     * @param string $default the default value for the property
     *
     * @return mixed
     * @throws Exception
     */
    protected function getSetting($key, $default = null)
    {
        $settings = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);

        return ArrayHelper::getValue($settings, $key, $default);
    }

    /**
     * Initialize columns selected for export
     */
    protected function initSelectedColumns()
    {
        if (!$this->_columnSelectorEnabled) {
            return;
        }
        $expCols = Yii::$app->request->post($this->exportColsParam, '');
        $this->selectedColumns = empty($expCols) ? array_keys($this->columnSelector) : Json::decode($expCols);
    }

    /**
     * Clear output buffers
     */
    protected function clearOutputBuffers()
    {
        if ($this->clearBuffers) {
            while (ob_get_level() > 0) {
                ob_end_clean();
            }
        } else {
            ob_end_clean();
        }
    }

    /**
     * Initialize column selector list
     */
    protected function initColumnSelector()
    {
        if (!$this->_columnSelectorEnabled) {
            return;
        }
        $selector = [];
        Html::addCssClass($this->columnSelectorOptions, ['btn', $this->getDefaultBtnCss(), 'dropdown-toggle']);
        $header = ArrayHelper::getValue($this->columnSelectorOptions, 'header', Yii::t('kvexport', 'Select Columns'));
        $this->columnSelectorOptions['header'] = (!isset($header) || $header === false) ? '' :
            '<li class="dropdown-header">' . $header . '</li><li class="kv-divider"></li>';
        $id = $this->options['id'] . '-cols';
        Html::addCssClass($this->columnSelectorMenuOptions, 'dropdown-menu kv-checkbox-list');
        $this->columnSelectorMenuOptions = array_replace_recursive(
            [
                'id'              => $id . '-list',
                'role'            => 'menu',
                'aria-labelledby' => $id,
            ],
            $this->columnSelectorMenuOptions
        );
        $dataToggle = 'data-' . ($this->isBs(5) ? 'bs-' : '') . 'toggle';
        $this->columnSelectorOptions = array_replace_recursive(
            [
                'id'            => $id,
                'icon'          => !$this->isBs(3) ? '<i class="fas fa-list"></i>' : '<i class="glyphicon glyphicon-list"></i>',
                'title'         => Yii::t('kvexport', 'Select columns to export'),
                'type'          => 'button',
                $dataToggle     => 'dropdown',
                'aria-haspopup' => 'true',
                'aria-expanded' => 'false',
            ],
            $this->columnSelectorOptions
        );
        foreach ($this->columns as $key => $column) {
            $selector[$key] = $this->getColumnLabel($key, $column);
        }
        $this->columnSelector = array_replace($selector, $this->columnSelector);
        if (!isset($this->selectedColumns)) {
            $keys = array_keys($this->columnSelector);
            $this->selectedColumns = array_combine($keys, $keys);
        }
    }

    /**
     * Fetches the column label
     *
     * @param integer $key
     * @param Column $column
     *
     * @return string
     */
    protected function getColumnLabel($key, $column)
    {
        if (is_int($key)) {
            $key++;
        }
        $label = Yii::t('kvexport', 'Column') . ' ' . $key;
        if (isset($column->label)) {
            $label = $column->label;
        } elseif (isset($column->header)) {
            $label = $column->header;
        } elseif (isset($column->attribute)) {
            $label = $this->getAttributeLabel($column->attribute);
        } elseif (!$column instanceof DataColumn) {
            $class = explode('\\', get_class($column));
            $label = Inflector::camel2words(end($class));
        }

        return trim(strip_tags(str_replace(['<br>', '<br/>'], ' ', $label)));
    }

    /**
     * Generates the attribute label
     *
     * @param string $attribute
     *
     * @return string
     */
    protected function getAttributeLabel($attribute)
    {
        /**
         * @var Model $model
         */
        $provider = $this->dataProvider;
        if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
            /** @noinspection PhpPossiblePolymorphicInvocationInspection */
            $modelClass = $provider->query->modelClass;
            $model = $modelClass::instance();

            return $model->getAttributeLabel($attribute);
        }
        if ($provider instanceof ActiveDataProvider && $provider->query instanceof QueryInterface) {
            return Inflector::camel2words($attribute);
        }
        $models = $provider->getModels();
        if (($model = reset($models)) instanceof Model) {
            return $model->getAttributeLabel($attribute);
        }
        return Inflector::camel2words($attribute);
    }

    /**
     * Sets the default export configuration
     * @throws InvalidConfigException|Exception
     */
    protected function setDefaultExportConfig()
    {
        $isFa = $this->fontAwesome;
        $notBs3 = !$this->isBs(3);
        $this->_defaultExportConfig = [
            self::FORMAT_CSV     => [
                'label'       => Yii::t('kvexport', 'CSV'),
                'icon'        => $notBs3 ? 'fas fa-file-code' : ($isFa ? 'fa fa-file-code-o' : 'glyphicon glyphicon-floppy-open'),
                'iconOptions' => ['class' => 'text-primary'],
                'linkOptions' => [],
                'options'     => ['title' => Yii::t('kvexport', 'Comma Separated Values')],
                'alertMsg'    => Yii::t('kvexport', 'The CSV export file will be generated for download.'),
                'mime'        => 'application/csv',
                'extension'   => 'csv',
                'writer'      => self::FORMAT_CSV,
                'delimiter'   => ",",
            ],
            self::FORMAT_TEXT    => [
                'label'       => Yii::t('kvexport', 'Text'),
                'icon'        => $notBs3 ? 'far fa-file-alt' : ($isFa ? 'fa fa-file-text-o' : 'glyphicon glyphicon-floppy-save'),
                'iconOptions' => ['class' => 'text-muted'],
                'linkOptions' => [],
                'options'     => ['title' => Yii::t('kvexport', 'Tab Delimited Text')],
                'alertMsg'    => Yii::t('kvexport', 'The TEXT export file will be generated for download.'),
                'mime'        => 'text/plain',
                'extension'   => 'txt',
                'writer'      => self::FORMAT_CSV,
                'delimiter'   => "\t",
            ],
            self::FORMAT_EXCEL_X => [
                'label'       => Yii::t('kvexport', 'Excel 2007+'),
                'icon'        => $notBs3 ? 'fas fa-file-excel' : ($isFa ? 'fa fa-file-excel-o' : 'glyphicon glyphicon-floppy-remove'),
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options'     => ['title' => Yii::t('kvexport', 'Microsoft Excel 2007+ (xlsx)')],
                'alertMsg'    => Yii::t('kvexport', 'The EXCEL 2007+ (xlsx) export file will be generated for download.'),
                'mime'        => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'extension'   => 'xlsx',
                'writer'      => self::FORMAT_EXCEL_X,
            ],
            self::FORMAT_ODS     => [
                'label'       => Yii::t('kvexport', 'OpenOffice'),
                'icon'        => $notBs3 ? 'fas fa-file-alt' : ($isFa ? 'fa fa-file-text' : 'glyphicon glyphicon-save'),
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options'     => ['title' => Yii::t('kvexport', 'OpenOffice (ods)')],
                'alertMsg'    => Yii::t('kvexport', 'The OPENOFFICE export file will be generated for download.'),
                'mime'        => 'application/vnd.oasis.opendocument.spreadsheet',
                'extension'   => 'ods',
                'writer'      => self::FORMAT_ODS,
            ],
        ];
    }

    /**
     * Registers client assets needed for Export Menu widget
     */
    protected function registerAssets()
    {
        $view = $this->getView();
        Dialog::widget($this->krajeeDialogSettings);
        ExportMenuAsset::register($view);
        $this->messages += [
            'allowPopups'      => Yii::t(
                'kvexport',
                'Disable any popup blockers in your browser to ensure proper download.'
            ),
            'confirmDownload'  => Yii::t('kvexport', 'Ok to proceed?'),
            'downloadProgress' => Yii::t('kvexport', 'Generating the export file. Please wait...'),
            'downloadComplete' => Yii::t(
                'kvexport',
                'Request submitted! You may safely close this dialog after saving your downloaded file.'
            ),
        ];
        $options = [
            'target'                 => $this->target,
            'formOptions'            => $this->exportFormOptions,
            'messages'               => $this->messages,
            'exportType'             => $this->_exportType,
            'colSelFlagParam'        => $this->colSelFlagParam,
            'colSelEnabled'          => $this->_columnSelectorEnabled ? 1 : 0,
            'exportRequestParam'     => $this->exportRequestParam,
            'exportTypeParam'        => $this->exportTypeParam,
            'exportColsParam'        => $this->exportColsParam,
            'exportFormHiddenInputs' => $this->exportFormHiddenInputs,
            'showConfirmAlert'       => $this->showConfirmAlert,
            'dialogLib'              => ArrayHelper::getValue($this->krajeeDialogSettings, 'libName', 'krajeeDialog'),
        ];
        if ($this->_columnSelectorEnabled) {
            $options['colSelId'] = $this->columnSelectorOptions['id'];
        }
        $options = Json::encode($options);
        $menu = 'kvexpmenu_' . hash('crc32', $options);
        $view->registerJs("var {$menu} = {$options};\n", View::POS_HEAD);
        $script = '';
        foreach ($this->exportConfig as $format => $setting) {
            if (!isset($setting) || $setting === false) {
                continue;
            }
            $id = $this->options['id'] . '-' . strtolower($format);
            $options = Json::encode(
                [
                    'settings' => new JsExpression($menu),
                    'alertMsg' => $setting['alertMsg'],
                ]
            );
            $script .= "jQuery('#{$id}').exportdata({$options});\n";
        }
        if ($this->_columnSelectorEnabled) {
            $id = $this->columnSelectorMenuOptions['id'];
            ExportColumnAsset::register($view);
            $script .= "jQuery('#{$id}').exportcolumns({});\n";
        }
        if (!empty($script) && isset($this->pjaxContainerId)) {
            $script .= "jQuery('#{$this->pjaxContainerId}').on('pjax:complete', function() {
                {$script}
            });\n";
        }
        $view->registerJs($script);
    }

    /**
     * Parses and returns the style options for `contentBefore` or `contentAfter`
     *
     * @param array $settings the settings to parse (for `contentBefore` or `contentAfter`)
     *
     * @return array
     * @throws Exception
     */
    protected function getStyleOpts($settings = [])
    {
        $styleOpts = ArrayHelper::getValue($settings, 'styleOptions', []);

        return ArrayHelper::getValue($styleOpts, $this->_exportType, []);
    }

    /**
     * Search all group-able columns
     */
    protected function findGroupedColumn()
    {
        foreach ($this->getVisibleColumns() as $key => $column) {
            $this->_groupedColumn[$key] = empty($column) || empty($column->group) ? null :
                ['firstLine' => -1, 'value' => null];
        }
        $this->_groupedColumn[] = null; //prevent the overflow
        $this->_groupedColumn[] = null; //prevent the overflow
    }

    /**
     * Validates a grouped row
     *
     * @param Model|array $model the data model
     * @param Model|array $nextModel the next data model
     * @param integer $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by
     * [[dataProvider]].
     */
    protected function checkGroupedRow($model, $nextModel, $key, $index)
    {
        $endCol = 0;
        /**
         * @var Column $column
         */
        foreach ($this->getVisibleColumns() as $column) {
            if ((isset($this->_groupedColumn[$endCol])) && (!is_null($this->_groupedColumn[$endCol]))) {
                $value = ($column->content === null) ? (method_exists($column, 'getDataCellValue') ?
                    $this->formatter->format($column->getDataCellValue($model, $key, $index), 'raw') :
                    $column->renderDataCell($model, $key, $index)) :
                    call_user_func($column->content, $model, $key, $index, $column);
                $nextValue = ($column->content === null) ? (method_exists($column, 'getDataCellValue') ?
                    $this->formatter->format($column->getDataCellValue($nextModel, $key, $index), 'raw') :
                    $column->renderDataCell($nextModel, $key, $index)) :
                    call_user_func($column->content, $nextModel, $key, $index, $column);
                if (is_null($this->_groupedColumn[$endCol]['value'])) {
                    $this->_groupedColumn[$endCol]['value'] = $value;
                    $this->_groupedColumn[$endCol]['firstLine'] = $index;
                }
                if ($this->_groupedColumn[$endCol]['value'] != $nextValue) {
                    $groupFooter = $column->groupFooter ?? null;
                    if ($groupFooter instanceof Closure) {
                        $groupFooter = call_user_func($groupFooter, $model, $key, $index, $this);
                    }
                    if (isset($groupFooter['content'])) {
                        $this->generateGroupedRow($groupFooter['content'], $endCol);
                    }
                    $this->_groupedColumn[$endCol]['firstLine'] = $index;
                }
                $this->_groupedColumn[$endCol]['value'] = $nextValue;
                $this->autoWidthColumns[$endCol] = max(strlen($nextValue), $this->autoWidthColumns[$endCol]);
            }
            $endCol++;
        }
    }

    /**
     * Generate a grouped row
     *
     * @param array $groupFooter footer row
     * @param integer $groupedCol the zero-based index of grouped column
     */
    protected function generateGroupedRow($groupFooter, $groupedCol)
    {
        $this->_groupedRow = [];
        $fLine = ArrayHelper::getValue($this->_groupedColumn[$groupedCol], 'firstLine', -1);
        $fLine = ($fLine == $this->_beginRow) ? $this->_beginRow + 1 : ($fLine + 3);
        $firstLine = ($this->_endRow == ($this->_beginRow + 2) && $fLine == 2) ? $this->_beginRow + 3 : $fLine;
        $endLine = $this->_endRow + 2;
        [$endLine, $firstLine] = ($endLine > $firstLine) ? [$endLine, $firstLine] : [$firstLine, $endLine];
        foreach ($this->getVisibleColumns() as $key => $column) {
            $value = $groupFooter[$key] ?? '';
            $groupedRange = self::columnName($key + 1) . $firstLine . ':' . self::columnName($key + 1) . $endLine;
            if (isset($column->group) && $column->group) {
                $this->mergeCells($key + 1, $firstLine, $key + 1, $endLine);
            }
            switch ($value) {
                case self::F_SUM:
                    $value = "=SUM($groupedRange)";
                    break;
                case self::F_COUNT:
                    $value = '=COUNTIF(' . $groupedRange . ',"*")';
                    break;
                case self::F_AVG:
                    $value = "=AVERAGE($groupedRange)";
                    break;
                case self::F_MAX:
                    $value = "=MAX($groupedRange)";
                    break;
                case self::F_MIN:
                    $value = "=MIN($groupedRange)";
                    break;
            }
            if ($value instanceof Closure) {
                $value = $value($groupedRange, $this);
            }
            $this->_groupedRow[] = !isset($value) || $value === '' ? '' : strip_tags($value);
        }
    }

    /**
     * Merge the given set of cells.
     * All column and row parameters are 1-indexed.
     *
     * @param int $topLeftColumn the leftmost column of the group to merge.
     * @param int $topLeftRow the top row of the group to merge.
     * @param int $bottomRightColumn The rightmost column of the group to merge.
     * @param int $bottomRightRow The bottom row of the group to merge.
     * @return void
     */
    protected function mergeCells($topLeftColumn, $topLeftRow, $bottomRightColumn, $bottomRightRow)
    {
        // mergeCells is only supported for Excel!
        if ($this->_objOpenspoutOptions instanceof Options) {
            $this->_objOpenspoutOptions->mergeCells($topLeftColumn - 1, $topLeftRow, $bottomRightColumn - 1, $bottomRightRow, $this->_objOpenspoutSheet->getIndex());
        }
    }

    /**
     * Cleans up the export file and current object instance
     *
     * @param string $file the file exported
     */
    protected function cleanup($file)
    {
        if ($this->stream || $this->deleteAfterSave) {
            @unlink($file);
        }
    }

    /**
     * Sanitizes file name
     * @param string $string
     * @return string
     */
    protected static function sanitize($string)
    {
        $reserved = array_merge(
            array_map('chr', range(0, 31)),
            ['?', '[', ']', '/', '\\', '=', '<', '>', ':', ';', ',', "'", '"', '&', '$', '#'],
            ['*', '(', ')', '|', '~', '`', '!', '{', '}', '%', '+', '’', '«', '»', '”', '“']
        );

        $string = str_replace($reserved, '-', trim($string));
        $string = preg_replace_callback(
            '/[^\x20-\x7f]/',
            function ($match) {
                return strtolower(str_replace('%', '', urlencode($match[0])));
            },
            $string
        );

        return trim($string, ' -');
    }
}
