<?php
defined('BASEPATH') or exit('No direct script access allowed');
 
//load Spout Library
require_once APPPATH . 'third_party/spout/src/Spout/Autoloader/autoload.php';
 
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
 
class Importexcel extends CI_Controller
{
 
    function __construct()
    {
        parent::__construct();
        //load model
        $this->load->model('app');
    }
 
    public function index()
    {
        //ketika button submit diklik
        if ($this->input->post('submit', TRUE) == 'upload') {
            $config['upload_path']      = './temp_doc/'; //siapkan path untuk upload file
            $config['allowed_types']    = 'xlsx|xls'; //siapkan format file
            $config['file_name']        = 'doc' . time(); //rename file yang diupload
 
            $this->load->library('upload', $config);
            
            if ($this->upload->do_upload('excel')) {
                //fetch data upload
                $file   = $this->upload->data();
                
                $reader = ReaderEntityFactory::createXLSXReader(); //buat xlsx reader
                // $reader->setShouldFormatDates(true); //will return formatted dates
                $reader->open('temp_doc/' . $file['file_name']); //open file xlsx yang baru saja diunggah
                
                echo "<pre>";
                //looping pembacaat sheet dalam file        
                foreach ($reader->getSheetIterator() as $sheet) {
                    $numRow = 1;
 
                    //siapkan variabel array kosong untuk menampung variabel array data
                    $save   = array();
 
                    //looping pembacaan row dalam sheet
                    foreach ($sheet->getRowIterator() as $row) {
 
                        if ($numRow > 1) {
                            //ambil cell
                            $cells = $row->getCells();
 
                            // $data = array(
                            //     'Id'              => $cells[0],
                            //     'Data_Sales_Id'     => $cells[1],
                            //     'Status'            => $cells[2],
                            //     'Join_Date'             => NULL,
                            //     'Resign_Date'            => $cells[4],
                            //     'Join_Date2'            => $cells[5],
                            //     'Resign_Date2'            => $cells[6],
                            //     'Reason'            => $cells[7],
                            //     'Foul_Type'            => $cells[8],
                            //     'Foul_Date'            => $cells[9],
                            //     'Foul_Rules'            => $cells[10],
                            //     'Foul_Notes'            => $cells[11],
                            // );

                            
                            // print_r($cells[4]);
                            var_dump($cells[3] != "" ? "ada" : "enggak");
                            //tambahkan array $data ke $save
                            // array_push($save, $data);
                        }
                        echo "<hr>";
                        // print_r($save);
 
                        $numRow++;
                    }
                    //simpan data ke database
                    die();
                    $this->app->simpan($save);
                    echo "</pre>";
                    //tutup spout reader
                    $reader->close();
 
                    //hapus file yang sudah diupload
                    unlink('temp_doc/' . $file['file_name']);
 
                    //tampilkan pesan success dan redirect ulang ke index controller import
                    echo    '<script type="text/javascript">
                              alert(\'Data berhasil disimpan\');
                              window.location.replace("' . base_url() . '");
                          </script>';
                }
            } else {
                echo "Error :" . $this->upload->display_errors(); //tampilkan pesan error jika file gagal diupload
            }
        }
 
        $this->load->view('import');
    }
}