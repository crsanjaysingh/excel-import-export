<?php

namespace App\Http\Controllers\Phpspreadsheet;

use App\Http\Controllers\Controller;
use DOMDocument;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class ExcelController extends Controller
{
    public function importExcel()
    {
        return view('Phpspreadsheet/excel-data-listing');
    }

    public function processExcel(Request $request)
    {
        $request->validate([
            'file' => 'required|mimes:xlsx,xls'
        ]);

        $file = $request->file('file');

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        return view('list', compact('data'));
    }
    public function htmlToExcel(Request $request){
            
        // $spreadsheet = new Spreadsheet();
        // $sheet = $spreadsheet->getActiveSheet();        
        // $dom = new DOMDocument();
        // $dom->loadHTML($request->htmlData);
        // $rows = $dom->getElementsByTagName('tr');
        
        // foreach ($rows as $row) {
        //     $cols = $row->getElementsByTagName('td');
        //     $rowData = [];
        //     foreach ($cols as $col) {
        //         $rowData[] = $col->nodeValue;
        //     }
        //     $sheet->fromArray([$rowData], null, 'A' . ($sheet->getHighestRow() + 1));
        // }
        // $writer = new Xlsx($spreadsheet);
        // $writer->save('output.xlsx'); 
        // $spreadsheet = new Spreadsheet();
        // $sheet = $spreadsheet->getActiveSheet();
        // $dom = new DOMDocument();
        // $dom->loadHTML($request->htmlData);
        // $rows = $dom->getElementsByTagName('tr');
        
        // foreach ($rows as $row) {
        //     $cols = $row->getElementsByTagName('td');
        //     $rowData = [];
        //     foreach ($cols as $col) {
        //         $rowData[] = $col->nodeValue;
        //     }
        //     $sheet->fromArray([$rowData], null, 'A' . ($sheet->getHighestRow() + 1));
        // }
        
        // $writer = new Xlsx($spreadsheet);
        
        // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // header('Content-Disposition: attachment;filename="output.xlsx"');
        // header('Cache-Control: max-age=0');
        
        // $writer->save('php://output');

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $dom = new DOMDocument();
        $dom->loadHTML($request->htmlData);
        $rows = $dom->getElementsByTagName('tr');
        
        foreach ($rows as $row) {
            $cols = $row->getElementsByTagName('td');
            $rowData = [];
            foreach ($cols as $col) {
                $rowData[] = $col->nodeValue;
            }
            $sheet->fromArray([$rowData]);
        }
        dd($sheet);
        $writer = new Xlsx($spreadsheet);
        
        $fileName = 'output.xlsx';
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');
        
        $writer->save('php://output');
        exit;
        
        
    }

}
