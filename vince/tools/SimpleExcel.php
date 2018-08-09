<?php
/**
 * Created by PhpStorm.
 * Powered by: Vince
 * Email: so_vince@outlook.com
 * Date: 2018/6/27
 * Time: 11:16
 */

namespace vince\tp_tools;


use think\Loader;

class SimpleExcel
{
    protected $objectExcel;
    protected $excelName;
    protected $letterArray=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];//excel列名
    protected $columnWidth = 20;
    protected $rowHeight;
    protected $sheetNum = 0;//当前sheet的编号
    protected $center = true;//水平垂直居中

    public function __construct()
    {
        Loader::import("com.PHPExcel.PHPExcel");
        $this->objectExcel = new \PHPExcel();
    }

    /**设置文件名
     * @param $name
     * @return $this
     */
    public function setExcelName($name){
        $this->excelName=$name;
        return $this;
    }

    /**获取文件名
     * @return string
     */
    public function getExcelName()
    {
        return $this->excelName ? $this->excelName : "新建的Excel-".time();
    }

    /** 设置列宽 ['A'=>20,'B'=>15]
     * @param $width
     * @return $this
     */
    public function setColumnWidth($width){
        if(is_integer($width)||is_array($width)){
            $this->columnWidth=$width;
        }
        return $this;
    }

    /**
     * 设置默认行高
     * Power: Mikkle
     * Email：776329498@qq.com
     * @param $height
     * @return $this
     */
    public function setRowHeight($height){
        if(is_numeric($height)){
            $this->rowHeight=$height;
        }
        return $this;
    }

    public function setCenter($bool){
        if(is_bool($bool)){
            $this->center = $bool;
        }
        return $this;
    }


    /**获取新的Sheet编号,初始为0
     * @return int
     */
    protected function getNewSheetNum(){
        $sheet_num=$this->sheetNum;
        $this->sheetNum=$sheet_num+1;
        return $sheet_num;
    }

    /**创建sheet,支持链式操作多个sheet
     * @param $sheet_title
     * @param $title   ['数据字段'=>'Excel列名']
     * @param array $data 数据
     * @param string $loop_field 子循环字段
     * @param array $merge_field 合并子循环中不包括的字段
     * @return $this
     */
    public function createSheet($sheet_title,$title,$data=[],$loop_field='',$merge_field = []){
        $letter_array = $this->letterArray;
        $field_title=array_values($title);
        $field_db = array_keys($title);

        $sheet_num = $this->getNewSheetNum();
        $objectExcel = $this->objectExcel;

        $objectExcel->createSheet($sheet_num);
        $objectExcel->setActiveSheetIndex($sheet_num);
        $objectExcel->getActiveSheet()->setTitle($sheet_title);

        $sheet=$objectExcel->getActiveSheet();

        //设置列宽
        foreach ($letter_array as $item){
            if(isset($this->columnWidth)){
                if(is_array($this->columnWidth)){
                    if(!empty($this->columnWidth[$item])) $sheet->getColumnDimension($item)->setWidth($this->columnWidth[$item]);
                    else $sheet->getColumnDimension($item)->setWidth(20);
                }else{
                    $sheet->getColumnDimension($item)->setWidth($this->columnWidth);
                }
            }
        }
        //设置默认行高
        if(!empty($this->rowHeight)){
            $sheet->getDefaultRowDimension()->setRowHeight($this->rowHeight);
        }
        //水平垂直居中
        if($this->center){
            foreach ($letter_array as $key=>$value){
                $sheet->getStyle($value)->getAlignment()->setHorizontal('center');
                $sheet->getStyle($value)->getAlignment()->setVertical('center');
            }
        }

        //标题填充
        foreach($field_title as $key=>$value){
            $sheet->setCellValue($letter_array[$key]."1",$value);
        }
        //数据填充
        if($data){
            if(empty($loop_field)){//是否有子循环，仅支持一个字段
                foreach ($data as $key=>$value){
                    foreach ($field_db as $k=>$v){
                        $sheet->setCellValue($letter_array[$k].($key+2),$value[$v]);
                    }
                }
            }else{
                $i = 2;
                foreach ($data as $key=>$value){
                    foreach ($value[$loop_field] as $k=>$v){
                        foreach ($field_db as $a=>$b){
                            if(in_array($b,$merge_field)) $sheet->setCellValue($letter_array[$a].$i,$value[$b]);
                            else $sheet->setCellValue($letter_array[$a].$i,$v[$b]);
                        }
                        $i++;
                    }

                    $c1 = count($value[$loop_field]);
                    $c2 = $c1 > 0 ? 1 : 0;

                    foreach ($field_db as $k=>$v){
                        if(in_array($v,$merge_field)){
                            $sheet->mergeCells($letter_array[$k].($i-$c1).':'.$letter_array[$k].($i-$c2));
                        }
                    }

                }
            }
        }

        return $this;
    }


    /**
     * 下载
     */
    public function downloadExcel(){
        ob_start();
        //最后通过浏览器输出
        $save_name=$this->getExcelName();
        $save_name = "$save_name.xls";
        header('Content-Type: application/vnd.ms-excel; charset=utf-8');
        header("Content-Disposition: attachment;filename=$save_name");
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($this->objectExcel, 'Excel5');
        $objWriter->save('php://output');

        ob_end_flush();//输出全部内容到浏览器
        die();
    }



}