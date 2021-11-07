<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Collection;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpWord\TemplateProcessor;
use Spatie\SimpleExcel\SimpleExcelReader;

class Controller extends BaseController
{
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    public function index() {
        return view('index');
    }

    public function importXlsx(Request $request) {   

        // $rows is an instance of Illuminate\Support\LazyCollection
        $rows = SimpleExcelReader::create('/Users/alanwilson/Documents/names.xlsx')->getRows();

        $created = collect();

        $rows->each(function(array $rowProperties) use ($created) {

            // dump($rowProperties);
           
            $phpWord = new TemplateProcessor('/Users/alanwilson/Downloads/Certificate.docx');
    
            $pathToSave = '/Users/alanwilson/Downloads/' . $rowProperties['name'] . '.docx' ;
    
                // $pathToSave = $name . '.docx';
            
            $phpWord->setValue('name', $rowProperties['name']);
            $phpWord->setValue('email', $rowProperties['email']);
            $phpWord->saveAs($pathToSave);

            $created->push($rowProperties);
    
        });

        return view('completed', compact('created'));


        $inputFileType = 'Xlsx';
        $inputFileName = '/Users/alanwilson/Documents/names.xlsx';

        /**  Create a new Reader of the type defined in $inputFileType  **/
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
        /**  Advise the Reader that we only want to load cell data  **/
        $reader->setReadDataOnly(true);
        /**  Load $inputFileName to a Spreadsheet Object  **/
        $spreadsheet = $reader->load($inputFileName);

        $phpExcel = new Spreadsheet();

        dump($phpExcel);


    }

    public function generateDocx(Request $request) {

        dump($request);

        // $phpWord = new PhpWord();
        

        // $section = $phpWord->addSection();

        $description = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod
tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam,
quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo
consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse
cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non
proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
        
        // $section->addText($description);

        $names = [
            'John Doe',
            'Jane Doe',
            'Kathryn Cooper'
        ];

        foreach ($names as $name) {
            $phpWord = new TemplateProcessor('/Users/alanwilson/Downloads/Certificate.docx');

            $pathToSave = '/Users/alanwilson/Downloads/' . $name . '.docx' ;

            // $pathToSave = $name . '.docx';
        
            $phpWord->setValue('name', $name);
            $phpWord->saveAs($pathToSave);

        }


        return 'Done';

        // $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        // try {
        //     $objWriter->save(storage_path('helloWorld.docx'));
        // } catch (\Exception $e) {
        // }

        // return response()->download(storage_path('helloWorld.docx'));

    }
}
