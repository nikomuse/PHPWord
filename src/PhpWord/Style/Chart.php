<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2014 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Style;

/**
 * Chart style
 *
 * @since 0.12.0
 */
class Chart extends AbstractStyle
{

    /**
     * Width (in EMU)
     *
     * @var int
     */
    private $width = 1000000;

    /**
     * Height (in EMU)
     *
     * @var int
     */
    private $height = 1000000;

    /**
     * Is 3D; applies to pie, bar, line, area
     *
     * @var bool
     */
    private $is3d = false;

    /**
     * Show labels for axis
     *
     * @var bool
     */
    private $showAxisLabels = false;
    /**
     * Show Gridlines for Y-Axis
     *
     * @var bool
     */
    private $gridY = false;
    /**
     * Show Gridlines for X-Axis
     *
     * @var bool
     */
    private $gridX = false;

    /**
     * Options
     *
     * @var array
     */
    private $options = array();
    
    /**
     * Legend
     *
     * @var boolean
     */
    private $legend = false;
    
    /**
     * Legend
     *
     * @var string
     */
    private $legend_position = '';
    
    /**
     * Data labels
     *
     * @var boolean
     */
    private $data_labels = false;
    
    /**
     * Data labels orientation
     *
     * @var string
     */
    private $data_labels_orientation = '';    
    
    /**
     * Major tick mark
     *
     * @var string
     */
    private $major_tick_mark = '';    
    
    /**
     * Legend top position
     * 
     * @var string
     */
    const LEGEND_POSITION_TOP = 't';
    
    /**
     * Legend right position
     * 
     * @var string
     */
    const LEGEND_POSITION_RIGHT = 'r';
    
    /**
     * Legend bottom position
     * 
     * @var string
     */
    const LEGEND_POSITION_BOTTOM = 'b';
    
    /**
     * Legend left position
     * 
     * @var string
     */
    const LEGEND_POSITION_LEFT = 'l';
    
    /**
     * Data labels normal orientation
     * 
     * @var string
     */
    const DATA_LABELS_ORIENTATION_NORMAL = 'dataLabelOrientationNormal';
    
    /**
     * Data labels 90° right orientation
     * 
     * @var string
     */
    const DATA_LABELS_ORIENTATION_RIGHT = 'dataLabelOrientationRight';
    
    /**
     * Data labels 90° left orientation
     * 
     * @var string
     */
    const DATA_LABELS_ORIENTATION_LEFT = 'dataLabelOrientationLeft';  
    
    /**
     * Tick mark : none
     * 
     * @var string
     */
    const TICK_MARK_NONE = 'none';
    
    /**
     * Tick mark in
     * 
     * @var string
     */
    const TICK_MARK_IN = 'in';
    
    /**
     * Tick mark out
     * 
     * @var string
     */
    const TICK_MARK_OUT = 'out';
    
    /**
     * Create a new instance
     *
     * @param array $style
     */
    public function __construct($style = array())
    {
        $this->setStyleByArray($style);
    }

    /**
     * Get width
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set width
     *
     * @param int $value
     * @return self
     */
    public function setWidth($value = null)
    {
        $this->width = $this->setIntVal($value, $this->width);

        return $this;
    }

    /**
     * Get height
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set height
     *
     * @param int $value
     * @return self
     */
    public function setHeight($value = null)
    {
        $this->height = $this->setIntVal($value, $this->height);

        return $this;
    }

    /**
     * Is 3D
     *
     * @return bool
     */
    public function is3d()
    {
        return $this->is3d;
    }

    /**
     * Set 3D
     *
     * @param bool $value
     * @return self
     */
    public function set3d($value = true)
    {
        $this->is3d = $this->setBoolVal($value, $this->is3d);

        return $this;
    }
    /**
     * Show labels for axis
     *
     * @return bool
     */
    public function showAxisLabels()
    {
        return $this->showAxisLabels;
    }
    /**
     * Set show Gridlines for Y-Axis
     *
     * @param bool $value
     * @return self
     */
    public function setShowAxisLabels($value = true)
    {
        $this->showAxisLabels = $this->setBoolVal($value, $this->showAxisLabels);
        return $this;
    }
    /**
     * Show Gridlines for Y-Axis
     *
     * @return bool
     */
    public function showGridY()
    {
        return $this->gridY;
    }
    /**
     * Set show Gridlines for Y-Axis
     *
     * @param bool $value
     * @return self
     */
    public function setShowGridY($value = true)
    {
        $this->gridY = $this->setBoolVal($value, $this->gridY);
        return $this;
    }
    /**
     * Show Gridlines for X-Axis
     *
     * @return bool
     */
    public function showGridX()
    {
        return $this->gridX;
    }
    /**
     * Set show Gridlines for X-Axis
     *
     * @param bool $value
     * @return self
     */
    public function setShowGridX($value = true)
    {
        $this->gridX = $this->setBoolVal($value, $this->gridX);
        return $this;
    }
    
    /**
     * Add a legend
     * 
     * @param string $position
     */
    public function addLegend($position='') {
        $this->legend_position = in_array($position, array(self::LEGEND_POSITION_TOP, self::LEGEND_POSITION_RIGHT, self::LEGEND_POSITION_BOTTOM, self::LEGEND_POSITION_LEFT)) ? $position : self::LEGEND_POSITION_BOTTOM;
        $this->legend = true;
    }

    /**
     * Check if this chart has a legend
     *
     * @return boolean
     */
    public function hasLegend()
    {
        return $this->legend;
    }

    /**
     * Get the legend position
     *
     * @return string
     */
    public function getLegendPosition()
    {
        return $this->legend_position;
    }
    
    /**
     * Add data labels
     * 
     * @param string $orientation
     */
    public function addDataLabels($orientation='') {
        $this->data_labels_orientation = in_array($orientation, array(self::DATA_LABELS_ORIENTATION_NORMAL, self::DATA_LABELS_ORIENTATION_RIGHT, self::DATA_LABELS_ORIENTATION_LEFT)) ? $orientation : self::DATA_LABELS_ORIENTATION_NORMAL;
        $this->data_labels = true;
    }

    /**
     * Check if this chart has data labels
     *
     * @return boolean
     */
    public function hasDataLabels()
    {
        return $this->data_labels;
    }

    /**
     * Get the data labels orientation
     *
     * @return string
     */
    public function getDataLabelsOrientation()
    {
        return $this->data_labels_orientation;
    }
    
    /**
     * Set major tick mark type
     * 
     * @param string $type
     */
    public function setMajorTickMark($type='') {
        $this->major_tick_mark = in_array($type, array(self::TICK_MARK_NONE, self::TICK_MARK_IN, self::TICK_MARK_OUT)) ? $type : self::TICK_MARK_NONE;
    }

    /**
     * Get major tick mark type
     *
     * @return string
     */
    public function getMajorTickMark()
    {
        return $this->major_tick_mark;
    }

    /**
     * Set the element options
     *
     * @return array
     */
    public function setOptions($options)
    {
        $this->options = $options;
    }

    /**
     * Set an element option
     *
     * @param string
     * @param string
     */
    public function setOption($key, $value)
    {
        $this->options[$key] = $value;
    }

    /**
     * Get an element option
     *
     * @param string $key
     * @return string
     */
    public function getOption($key)
    {
        return isset($this->options[$key]) ? $this->options[$key] : '';
    }
}
