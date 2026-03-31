/*!
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @version   2.0.0
 *
 * Export Columns Selector Validation Module.
 *
 */
(function ($) {
    "use strict";

    const EXPORT_COLUMNS_SELECTOR_CHECKBOX = 'input[name="export_columns_selector[]"]';

    const ExportColumns = function (element, options) {
        const self = this;
        self.$element = $(element);
        self.options = options;
        self.listen();
    };

    ExportColumns.prototype = {
        constructor: ExportColumns,
        listen: function () {
            const self = this;
            const $el = self.$element;
            const $tog = $el.find('input[name="export_columns_toggle"]');
            $el.off('click').on('click', function (e) {
                e.stopPropagation();
            });
            $tog.off('change').on('change', function () {
                const checked = $tog.is(':checked');
                $el.find(EXPORT_COLUMNS_SELECTOR_CHECKBOX + ':not([disabled])').prop('checked', checked);
            });
            let otherChildren;
            const getElementBeforeAndAfterSortedElement = (offset) => {
                let firstElemAboveSorted = null;
                let firstOffsetBelow = null;
                let firstElemBelowSorted = null;
                let firstOffsetAbove = null;
                otherChildren.each(function () {
                    const otherRow = $(this);
                    const elemOffset = otherRow.offset().top;
                    if (elemOffset < offset && (firstOffsetBelow === null || firstOffsetBelow < elemOffset)){
                        firstElemAboveSorted = otherRow;
                        firstOffsetBelow = elemOffset;
                        return;
                    }
                    if (elemOffset >= offset && (firstOffsetAbove === null || firstOffsetAbove > elemOffset)) {
                        firstElemBelowSorted = otherRow;
                        firstOffsetAbove = elemOffset;
                    }
                });
                return {firstElemAboveSorted, firstElemBelowSorted};
            }
            $el.find('li').has(EXPORT_COLUMNS_SELECTOR_CHECKBOX)
                .draggable({
                    axis: 'y',
                    handle: 'label',
                    start: function (e, ui) {
                        const sortingRow = $(this);
                        sortingRow.css('z-index', 50);
                        const sortingId = sortingRow.find(EXPORT_COLUMNS_SELECTOR_CHECKBOX).data('key');
                        otherChildren = $el.find('li').has(EXPORT_COLUMNS_SELECTOR_CHECKBOX + ':not([data-key="' + sortingId + '"])');
                    },
                    drag: function (e, ui) {
                        const sortingRow = $(this);
                        $el.find('.highlight-sorting-below, .highlight-sorting-above').removeClass(['highlight-sorting-below', 'highlight-sorting-above']);
                        const offset = sortingRow.offset().top;
                        const {firstElemBelowSorted, firstElemAboveSorted} = getElementBeforeAndAfterSortedElement(offset);
                        if (firstElemAboveSorted !== null) {
                            firstElemAboveSorted.addClass('highlight-sorting-below');
                        }
                        if (firstElemBelowSorted !== null) {
                            firstElemBelowSorted.addClass('highlight-sorting-above');
                        }
                    },
                    stop: function () {
                        const sortingRow = $(this);
                        const offset = sortingRow.offset().top;
                        const {firstElemBelowSorted, firstElemAboveSorted} = getElementBeforeAndAfterSortedElement(offset);
                        if (firstElemBelowSorted !== null) {
                            sortingRow.insertBefore(firstElemBelowSorted);
                        } else if (firstElemAboveSorted !== null) {
                            sortingRow.insertAfter(firstElemAboveSorted);
                        }
                        $el.find('.highlight-sorting-below, .highlight-sorting-above').removeClass(['highlight-sorting-below', 'highlight-sorting-above']);
                        sortingRow.css('top', 0);
                    }
                });
        },
        getSortedIds: function () {
            return $el.find('input[name="export_columns_selector[]"]')
                .toArray()
                .sort((a, b) => {
                    const offsetA = $(a).offset().top;
                    const offsetB = $(b).offset().top;
                    return offsetA - offsetB;
                })
                .map(elem => $(elem).data('key'));
        },
    };

    //ExportColumns plugin definition
    $.fn.exportcolumns = function (option) {
        const args = Array.apply(null, arguments);
        args.shift();
        return this.each(function () {
            const $this = $(this);
            let data = $this.data('exportcolumns');
            let options = typeof option === 'object' && option;

            if (!data) {
                data = new ExportColumns(this, $.extend({}, $.fn.exportcolumns.defaults, options, $(this).data()))
                $this.data('exportcolumns', data);
            }

            if (typeof option === 'string') {
                data[option].apply(data, args);
            }
        });
    };

    $.fn.exportcolumns.defaults = {};
})(window.jQuery);