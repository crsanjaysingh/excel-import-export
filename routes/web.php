<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\Phpspreadsheet\ExcelController;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});


Route::get('/', [ExcelController::class, 'importExcel'])->name('import.excel');
Route::post('/process-excel', [ExcelController::class, 'processExcel'])->name('process.excel');
Route::post('/html-to-excel', [ExcelController::class, 'htmlToExcel'])->name('html.to.excel');


