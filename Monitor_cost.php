<?php
/* error_reporting(-1);
ini_set('display_errors', 1); */
class Monitor_cost {

    protected $CI;
    public $root_menu_id;
    public $current_menu_id;
    public $class_url;
    public $browse_url;
    public $class_tb;
    public $primary_key;
    public $browse_array = array();
    public $user_languages;

    /**
     * JOIN
     *
     * @var array
     * */
    public $_joins = array();

    /**
     * Extra Joins
     *
     * @var string
     **/
    public $_extra_joins = NULL;

    /**
     * Where
     *
     * @var array
     * */
    public $_where = array();

    /**
     * Custom Where
     *
     * @var array
     * */
    public $_custom_where = NULL;

    /**
     * Like
     *
     * @var array
     * */
    public $_like = array();

    /**
     * Or Like
     *
     * @var array
     * */
    public $_or_like = array();

    /**
     * Order By
     *
     * @var string
     * */
    public $_order_by = NULL;
    public $_group_by = NULL;

    /**
     * Custom pagination settings
     */
    public $pagination_config;
    //OFF6985
    public $xlsColumns = array();

    /**
     * Constructor
     */
    public function __construct() {
        $this->CI = & get_instance();
        $this->CI->load->helper('form');
        $this->CI->load->library('form_validation');
        $this->CI->load->library('auth_lib');
        $this->CI->load->library('pagination');
        //$this->CI->load->library('tab/monitor_lib');

        // init class variables
        $this->root_menu_id = 140;
        $this->current_menu_id = 232;

        $this->class_url = 'tab/monitor_costs';   // IMPORTANTE: senza slash finale!!!
        $this->browse_url = 'tab/monitor_costs/index';  // IMPORTANTE: senza slash finale!!!

        $this->class_tb = $this->CI->db->tb_article_cost;  // tabella principale della procedura
        $this->primary_key = 'id_article_cost'; // nome del campo chiave primaria sul db
        // array per definire le colonne della browse

        /*
        Verrà realizzato un modello “COSTO”, simile a quello già esistente per il monitor ordine cliente “OC”, in cui sarà possibile selezionare quali tra le seguenti colonne aggiungere nella browse:

        Costo dei componenti di acquisto 	tab_article_cost.component_acq_cost -> ok
        costo dei componenti di produzione	tab_article_cost.component_prd_cost -> ok
        Costo di assemblaggio			tab_article_cost.assembly_cost -> ok
        costi indiretti				tab_article_cost_detail.indirect_cost -> ok
        totale imposte 				tab_article_cost_detail.taxes -> ok
        totale costo (pre imposte)		tab_article_cost_detail.cost -> ok
        percentuale di ricarico pre imposte	DA calcolare o da salvare?
        percentuale di ricarico (comprese di imposte) DA calcolare o da salvare?
        data ultimo aggiornamento
        utente che ha eseguito l’aggiornamento

        /*************************//*
        Filtri
        I filtri di ingresso sono:
        Codice del listino da utilizzare per il calcolo del costo, proposto di default “VENDITA”. Dato obbligatorio tra i codici della tabella tab_price_list.
        Anno: Anno di riferimento per le tabelle dei costi, proposto anno corrente  - 1
        Articolo: filtro non obbligatorio, visualizzare tutti gli articoli che iniziano per il codice digitato
        Descrizione: filtro non obbligatorio, visualizzare tutti gli articoli la cui descrizione contiene il codice digitato
        Nella browse verranno visualizzati tutti i record della tabella tab_article_cost e tab_article_cost_detail che soddisfano i filtri iniziali con version = 0. Ogni riga avrà la possibilità di visualizzare o modificare il proprio dettaglio e di stamparlo in pdf. Deve essere possibile modificare i filtri iniziali. La browse deve essere esportabile in xls. Verrà data la possibilità di effettuare la scelta multipla sui codici articolo per eseguire il ricalcolo del costo tramite un’azione dedicata.
         */


        //articolo
        //descrizione
        $this->browse_array[] = array(
            'db_table'			=> $this->class_tb,
            'db_field_name'		=> "(CONCAT('[', " . "get_item_code({$this->class_tb}.fk_item_table,{$this->class_tb}.fk_item_id)" . ", '] - '," . "get_item_description_only({$this->class_tb}.fk_item_table,{$this->class_tb}.fk_item_id)" .  "))",
            'alias'				=> 'item',
            'column_label'		=> lang('item'),
            //'number_format'		=> array(2,",",""),
            //'extra_class'		=> 'text-right',
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
        );

        //totale costo (compreso di imposte)	tab_article_cost_detail.final_cost
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            'db_field_name'		=> "final_cost",
            'column_label'		=> lang('final_cost'),
            'number_format'		=> array(2,",",""),
            'extra_class'		=> 'text-right',
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
        );

        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            'db_field_name'		=> "sale_price",
            'column_label'		=> lang('sale_price'),
            'number_format'		=> array(2,",",""),
            'extra_class'		=> 'text-right',
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
        );

        //margine				(Prezzo di vendita - Totale costo) / prezzo di vendita %
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            'db_field_name'		=> "(" . $this->CI->db->tb_article_cost_detail . ".sale_price - " . $this->CI->db->tb_article_cost_detail . ".final_cost) / " . $this->CI->db->tb_article_cost_detail . ".sale_price",
            'column_label'		=> lang('margin'),
            'number_format'		=> array(2,",",""),
            'extra_class'		=> 'text-right',
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'			=> FALSE,
        );

        /************/

        //Costo dei componenti di acquisto 	tab_article_cost.component_acq_cost
        $this->browse_array[] = array(
            'db_table'			=> $this->class_tb,
            'db_field_name'		=> "component_acq_cost",
            //'column_label'		=> lang('purchase_component_cost'),
            'alias'             => "attribute_28",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //costo dei componenti di produzione	tab_article_cost.component_prd_cost
        $this->browse_array[] = array(
            'db_table'			=> $this->class_tb,
            'db_field_name'		=> "component_prd_cost",
            //'column_label'		=> lang('component_prd_cost'),
            'alias'             => "attribute_29",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //Costo di assemblaggio			tab_article_cost.assembly_cost
        $this->browse_array[] = array(
            'db_table'			=> $this->class_tb,
            'db_field_name'		=> "assembly_cost",
            //'column_label'		=> lang('assembly_cost'),
            'alias'             => "attribute_30",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //costi indiretti				tab_article_cost_detail.indirect_cost
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            'db_field_name'		=> "indirect_cost",
            //'column_label'		=> lang('indirect_cost'),
            'alias'             => "attribute_31",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //totale imposte 				tab_article_cost_detail.taxes
        $this->browse_array[] = array(
            'db_table'			=> $this->class_tb,
            'db_field_name'		=> "taxes",
            //'column_label'		=> lang('taxes'),
            'alias'             => "attribute_32",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //totale costo (pre imposte)		tab_article_cost_detail.cost
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            'db_field_name'		=> "cost",
            //'column_label'		=> lang('cost'),
            'alias'             => "attribute_33",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //data ultimo aggiornamento		tab_article_cost_detail.date_modify
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost,
            'db_field_name'		=> $this->CI->db->tb_article_cost.".user_modify",
            //'column_label'		=> lang('cost'),
            'alias'             => "attribute_34",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //utente ultimo aggiornamento		tab_article_cost_detail.user_modify
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost,
            'db_field_name'		=> $this->CI->db->tb_article_cost.".date_modify",
            //'column_label'		=> lang('cost'),
            'alias'             => "attribute_35",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //percentuale di ricarico no taxes
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            //'db_field_name'		=> ((($this->CI->db->tb_article_cost_detail.'.sale_price' - $this->CI->db->tb_article_cost_detail.'.cost') / $this->CI->db->tb_article_cost_detail.'.sale_price') * 100),
            'db_field_name'		=> "CASE WHEN {$this->CI->db->tb_article_cost_detail}.sale_price IS NULL OR {$this->CI->db->tb_article_cost_detail}.sale_price = 0 THEN 0 ELSE ((({$this->CI->db->tb_article_cost_detail}.sale_price - {$this->CI->db->tb_article_cost_detail}.cost) / {$this->CI->db->tb_article_cost_detail}.sale_price) * 100) END",
            //'column_label'		=> lang('cost'),
            'alias'             => "attribute_36",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );

        //percentuale di ricarico
        $this->browse_array[] = array(
            'db_table'			=> $this->CI->db->tb_article_cost_detail,
            //'db_field_name'		=> ((($this->CI->db->tb_article_cost_detail.'.sale_price' - $this->CI->db->tb_article_cost_detail.'.final_cost') / $this->CI->db->tb_article_cost_detail.'.sale_price') * 100),
            'db_field_name'		=> "CASE WHEN {$this->CI->db->tb_article_cost_detail}.sale_price IS NULL OR {$this->CI->db->tb_article_cost_detail}.sale_price = 0 THEN 0 ELSE (((ROUND({$this->CI->db->tb_article_cost_detail}.sale_price, 2) - ROUND({$this->CI->db->tb_article_cost_detail}.final_cost, 2)) / ROUND({$this->CI->db->tb_article_cost_detail}.final_cost, 2)) * 100) END",
            //'column_label'		=> lang('cost'),
            'alias'             => "attribute_37",
            'sortable'			=> TRUE,
            'no_table_prefix'	=> TRUE,
            'visible'          => FALSE
        );


        $this->user_languages = $this->CI->utility->get_user_languages();

        $this->_joins = array(
            $this->CI->db->tb_article_cost_detail => $this->class_tb . '.id_article_cost = ' . $this->CI->db->tb_article_cost_detail . '.fk_article_cost_id',
            $this->CI->db->tb_price_list => $this->CI->db->tb_article_cost_detail. '.fk_price_list_id = ' . $this->CI->db->tb_price_list . '.id_price_list',
        );

        $version = 0;
        $this->_where = array($this->class_tb . '.version'  => $version);

        // set default error delimiter
        $this->CI->form_validation->set_error_delimiters('<span class="help-block">', '</span>');

        $this->_group_by = $this->class_tb. '.id_article_cost';

        // set custom config pagination
        $this->pagination_config['base_url'] = site_url($this->browse_url);
        $this->pagination_config['uri_segment'] = 5;

        $this->pagination_config['per_page'] = $this->CI->config->item('records_per_page');
    }

    public function get_records($page = 0, $all = FALSE) {

        $from = 0;
        if ($page > 0)
            $from = ($page - 1) * $this->pagination_config['per_page'];

        $this->CI->db->where($this->_where);
        if (!is_null($this->_custom_where)) {
            foreach ($this->_custom_where AS $condition) {
                $this->CI->db->where($condition);
            }
        }

        $this->CI->db->like($this->_like);
        $this->CI->db->or_like($this->_or_like);
        if (!$all)
            $this->CI->db->limit($this->pagination_config['per_page'], $from);

        $select_fields = array($this->class_tb . '.' . $this->primary_key);
        $joins = array();
        for ($i = 0; $i < count($this->browse_array); $i++) {
            if (isset($this->browse_array[$i]['alias']) && $this->browse_array[$i]['alias'] != '') {

                if (isset($this->browse_array[$i]['is_multilang']) && $this->browse_array[$i]['is_multilang']) {

                    if ($this->CI->session->userdata('language_code') != '') {
                        $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['alias'] : $this->browse_array[$i]['db_table'] . '.' . $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['alias'];
                    } else {
                        $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? 'ml_' . $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->browse_array[$i]['alias'] : $this->browse_array[$i]['db_table'] . '.ml_' . $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->browse_array[$i]['alias'];
                    }
                } else {
                    $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->browse_array[$i]['alias'] : $this->browse_array[$i]['db_table'] . '.' . $this->browse_array[$i]['db_field_name'] . ' AS ' . $this->browse_array[$i]['alias'];
                }
            } else {

                if (isset($this->browse_array[$i]['is_multilang']) && $this->browse_array[$i]['is_multilang']) {

                    if ($this->CI->session->userdata('language_code') != '') {
                        $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['db_field_name'] : $this->browse_array[$i]['db_table'] . '.' . $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$i]['db_field_name'];
                    } else {
                        $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? 'ml_' . $this->browse_array[$i]['db_field_name'] : $this->browse_array[$i]['db_table'] . '.ml_' . $this->browse_array[$i]['db_field_name'];
                    }
                } else {

                    $select_fields[] = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->browse_array[$i]['db_field_name'] : $this->browse_array[$i]['db_table'] . '.' . $this->browse_array[$i]['db_field_name'];
                }
            }

            if (isset($this->browse_array[$i]['join_on'])) {
                $joins[$this->browse_array[$i]['db_table']] = $this->browse_array[$i]['join_on'];
            }
        }


        if (count($this->_joins) > 0) {
            foreach ($this->_joins AS $key => $value) {
                if (array_key_exists($key, $joins)) {
                    continue;
                }
                if (is_array($value)) {
                    if ($value[1] == TRUE) {
                        $this->CI->db->join($key, $value[0]);
                    } else {
                        $this->CI->db->join($key, $value[0], 'LEFT');
                    }
                } else {
                    $this->CI->db->join($key, $value, 'LEFT');
                }
            }
        }

        if (count($joins) > 0) {
            foreach ($joins AS $key => $value) {
                $this->CI->db->join($key, $value, 'LEFT');
            }
        }

        $this->CI->db->select(implode(',', $select_fields), FALSE);

        // imposto ordinamento
        if ($this->CI->session->userdata($this->class_tb . '_current_sort') && $this->CI->session->userdata($this->class_tb . '_current_sort') != '') {
            $data_sort = explode('|', $this->CI->session->userdata($this->class_tb . '_current_sort'));

            if (isset($this->browse_array[$data_sort[0]])) {

                if (isset($this->browse_array[$data_sort[0]]['alias'])) {

                    if (isset($this->browse_array[$data_sort[0]]['is_multilang']) && $this->browse_array[$data_sort[0]]['is_multilang']) {

                        if ($this->CI->session->userdata('language_code') != '') {
                            $this->_order_by = $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$data_sort[0]]['alias'] . ' ' . $data_sort[1];
                        } else {
                            $this->_order_by = $this->browse_array[$data_sort[0]]['alias'] . ' ' . $data_sort[1];
                        }
                    } else {
                        $this->_order_by = $this->browse_array[$data_sort[0]]['alias'] . ' ' . $data_sort[1];
                    }
                } else {

                    if (isset($this->browse_array[$data_sort[0]]['is_multilang']) && $this->browse_array[$data_sort[0]]['is_multilang']) {

                        if ($this->CI->session->userdata('language_code') != '') {
                            $this->_order_by = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1] : $this->browse_array[$data_sort[0]]['db_table'] . '.' . $this->CI->session->userdata('language_code') . '_ml_' . $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1];
                        } else {
                            $this->_order_by = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? 'ml_' . $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1] : $this->browse_array[$data_sort[0]]['db_table'] . '.ml_' . $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1];
                        }
                    } else {
                        $this->_order_by = (isset($this->browse_array[$i]['no_table_prefix']) && $this->browse_array[$i]['no_table_prefix']) ? $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1] : $this->browse_array[$data_sort[0]]['db_table'] . '.' . $this->browse_array[$data_sort[0]]['db_field_name'] . ' ' . $data_sort[1];
                    }
                }
            }
        }

        $this->CI->db->group_by($this->_group_by);
        $this->CI->db->order_by($this->_order_by);

        $query = $this->CI->db->get($this->class_tb);
        //error_log(print_r($this->CI->db->last_query(), true));
        if ($query->num_rows() > 0) {
            foreach ($query->result() as $row) {
                $data[] = $row;
            }

            return $data;
        }

        return false;
    }

    public function count_records() {

        $this->CI->db->select($this->primary_key);
        $this->CI->db->where($this->_where);
        if (!is_null($this->_custom_where)) {
            foreach ($this->_custom_where AS $condition) {
                $this->CI->db->where($condition);
            }
        }
        $this->CI->db->like($this->_like);
        $this->CI->db->or_like($this->_or_like);

        if (count($this->_joins) > 0) {
            foreach ($this->_joins AS $key => $value) {
                if (is_array($value)) {
                    if ($value[1] == TRUE) {
                        $this->CI->db->join($key, $value[0]);
                    } else {
                        $this->CI->db->join($key, $value[0], 'LEFT');
                    }
                } else {
                    $this->CI->db->join($key, $value, 'LEFT');
                }
            }
        }

        $this->CI->db->group_by($this->_group_by);

        return $this->CI->db->get($this->class_tb)->num_rows();
    }

    public function export() {

        $records = $this->get_records(0, TRUE);

        if ($records) {
            //load our new PHPExcel library
            $this->CI->load->library('excel');

            //activate worksheet number 1
            $this->CI->excel->setActiveSheetIndex(0);

            //name the worksheet
            $this->CI->excel->getActiveSheet()->setTitle(strtoupper(lang('situation_at')) . date('d-m-Y H.i'));

            // imposta il layout di stampa orizzontale
            $this->CI->utility->set_landscape_excel_layout($this->CI->excel);

            // Field names in the first row
            $this->CI->utility->print_excel_head($this->browse_array, $this->CI->excel);

            // custom columns
            $index = 5;
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($index, 1, strip_tags(lang('transport_number')));
            $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($index, 1)->setAutoSize(true);
            $index++;
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($index, 1, strip_tags(lang('colli')));
            $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($index, 1)->setAutoSize(true);

            // Fetching the table data
            $row = 2;
            foreach ($records as $data) {
                $this->CI->utility->print_excel_row($this->browse_array, $this->CI->excel, $data, $row);

                $num_transport = $this->CI->utility->get_num_transport_by_id_pallet($data->id_odl_pallet);
                $transp = '';
                foreach ($num_transport as $elm){
                    $transp .= $elm->num_transport.' ' ;
                }

                $colli = $this->CI->utility->get_records($this->CI->db->tb_odl_pallet_detail, array('fk_odl_pallet_id' => $data->id_odl_pallet) , 'from_pack, to_pack');
                $num_colli = '';
                foreach ($colli as $elm){
                    $num_colli .= $elm->from_pack . '-'.$elm->to_pack .'| ';
                }

                $index = 5;
                $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($index, $row, strip_tags($transp));
                $index++;
                $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($index, $row, strip_tags($num_colli));

                $row++;
            }

            $this->CI->excel->setActiveSheetIndex(0);
            $filename = str_replace(' ', '_', lang('menu_monitor_load_merchandise')) . '_' . date('dmY'); //save our workbook as this file name
            @ob_end_clean();
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
            header("Content-Type: application/force-download");
            header("Content-Type: application/octet-stream");
            header("Content-Type: application/download");
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header("Content-Disposition: attachment;filename=$filename.xlsx");
            header("Content-Transfer-Encoding: binary ");
            $objWriter = PHPExcel_IOFactory::createWriter($this->CI->excel, 'Excel2007');
            $objWriter->save('php://output');
        }
    }

    /**
     *
     */
    public function get_production_components_cost($id_article_cost, $year) {
        //tab_component_cost NON tipo "ACQ"
        //testa
        /*
        codice
        descrizione
        quantità di impiego				(da tab_bom_cost)
        peso 						(da tab_component)
        Costo materiale					tab_component_cost.material_cost
        tempo impiegato per realizzare un pezzo 	tab_component_cost.theoric_time
        costo orario del reparto (€/ora)			da configurazione
        Costo Produzione				= tempo * costo orario * quantità di impiego
        costo del componente				= costo produzione + costo materiale
        */

        //riga dettaglio
        /*
        Codice
        Descrizione
        quantità di impiego
        Costo di acquisto				tab_component_cost.cost
        */
        $production_components = array();
        $totals = array(
            'hour_cost' => 0,
            'production_cost' => 0,
            'component_cost' => 0,
            'material_cost' => 0
        );

        //alla fine riga con somma del totale
        $article_cost_record = $this->CI->utility->get_record_by_id($this->primary_key, $id_article_cost, $this->class_tb);
        //$article_cost_detail_record = $this->utility->get_record_by_id('fk_article_cost_id', $id_article_cost, $this->CI->db->tb_article_cost_detail);
        //error_log(print_r($article_cost_record, true));
        //con questi due vado a leggere la component_cost e la component con il codice articolo
        $article_table = $article_cost_record->fk_item_table;
        $article_id = $article_cost_record->fk_item_id;

        //prendo da tab bom cost con articolo
        $tab_bom_cost_records = $this->CI->utility->get_records($this->CI->db->tb_bom_cost, array('fk_parent_table' => $article_table, 'fk_parent_id' => $article_id));
        //error_log(print_r($tab_bom_cost_records, true));

        if(!empty($tab_bom_cost_records)) {
            //
            foreach ($tab_bom_cost_records as $tab_bom_cost_record) {
                //ciclo gli elementi legati all'articolo
                $component_child_table = $tab_bom_cost_record->fk_child_table;
                $component_child_id = $tab_bom_cost_record->fk_child_id;
                //per ognuno che rientra in tab_component_cost
                //vado a prendere il suo record in component cost
                //$component_cost_record = $this->CI->utility->get_record_by_id('id_component', $component_child_id, $component_child_table);
                $component_cost_record = array();
                $component_cost_records = $this->CI->utility->get_records($this->CI->db->tb_component_cost, array('fk_item_table' => $component_child_table, 'fk_item_id' => $component_child_id, 'year' => $year));
                //error_log(print_r($component_cost_records, true));
                if(!empty($component_cost_records)) {
                    $component_cost_record = $component_cost_records[0];
                }

                if(!empty($component_cost_record)) {
                    //se non è di acquisto allora lo prendo
                    if (!$component_cost_record->is_purchase) {
                        $details = array();
                        //
                        //cerco i suoi componenti di acquisto
                        $tab_bom_cost_purchase_component_records = $this->CI->utility->get_records($this->CI->db->tb_bom_cost, array('fk_parent_table' => $component_child_table, 'fk_parent_id' => $component_child_id));
                        if(!empty($tab_bom_cost_purchase_component_records)) {
                            foreach($tab_bom_cost_purchase_component_records as $tab_bom_cost_purchase_component_record) {
                                //$purchase_component_child_table = $tab_bom_cost_purchase_component_record->fk_child_table;
                                //$purchase_component_child_id = $tab_bom_cost_purchase_component_record->fk_child_id;
                                $purchase_component_table = $tab_bom_cost_purchase_component_record->fk_child_table;
                                $purchase_component_id = $tab_bom_cost_purchase_component_record->fk_child_id;
                                //
                                //$purchase_component_cost_record = $this->CI->utility->get_record_by_id('id_component', $purchase_component_child_id, $purchase_component_child_table);
                                $purchase_component_cost_record = array();
                                //$purchase_component_cost_records = $this->CI->utility->get_records($this->CI->db->tb_component_cost, array('fk_item_table' => $purchase_component_child_table, 'fk_item_id' => $purchase_component_child_id));
                                $purchase_component_cost_records = $this->CI->utility->get_records($this->CI->db->tb_component_cost, array('fk_item_table' => $purchase_component_table, 'fk_item_id' => $purchase_component_id, 'year' => $year));
                                if(!empty($purchase_component_cost_records)) {
                                    $purchase_component_cost_record = $purchase_component_cost_records[0];
                                }
                                //error_log(print_r($tab_bom_cost_purchase_component_record, true));

                                //error_log(print_r($purchase_component_cost_record, true));
                                if (!empty($purchase_component_cost_record)) {
                                    if (1 || $purchase_component_cost_record->is_purchase) {
                                        $detail_component = array();
                                        $detail_code = '';
                                        $detail_description = '';
                                        //
                                        if($purchase_component_table == 'tab_component') { //usare config db
                                            $detail_component = $this->CI->utility->get_record_by_id('id_component', $purchase_component_id, $purchase_component_table);
                                        } else if ($purchase_component_table == 'tab_packaging') { //usare config db
                                            $detail_component = $this->CI->utility->get_record_by_id('id_packaging', $purchase_component_id, $purchase_component_table);
                                        }
                                        //
                                        if(!empty($detail_component)) {
                                            $detail_code = $detail_component->code;
                                            //parametrizzare in base alla lingua di sistema
                                            $detail_description = $detail_component->it_ml_description;
                                        }
                                        //pusho i figli -> dettaglio componenti di acquisto
                                        $comp_details = array(
                                            'code'              => $detail_code,
                                            'description'       => $detail_description,
                                            'is_purchase'       => $purchase_component_cost_record->is_purchase,
                                            'qty'               => $tab_bom_cost_purchase_component_record->qty,
                                            'weight'            => $purchase_component_cost_record->weight,
                                            'cost'              => $purchase_component_cost_record->cost,
                                            'effect_percentage' => $purchase_component_cost_record->effect_percentage,
                                            'theoric_time'      => $purchase_component_cost_record->theoric_time,
                                            'hour_cost'         => $purchase_component_cost_record->hour_cost,
                                            'production_cost'   => $purchase_component_cost_record->production_cost,
                                        );
                                        $this->recursive_bom($comp_details, $purchase_component_id, $purchase_component_table, $year);
                                        array_push($details, $comp_details);

                                    }
                                }
                            }
                        }

                        $component = array();
                        $code = '';
                        $description = '';
                        if($component_child_table == 'tab_component') {
                            $component = $this->CI->utility->get_record_by_id('id_component', $component_child_id, $component_child_table);
                        } else if ($component_child_table == 'tab_kit' /*'tab_packaging'*/) {
                            $component = $this->CI->utility->get_record_by_id('id_kit' /*'id_packaging'*/, $component_child_id, $component_child_table);
                        }
                        //
                        if(!empty($component)) {
                            $code = $component->code;
                            //parametrizzare in base alla lingua di sistema
                            $description = $component->it_ml_description;
                        }
                        //$production_cost = ($component_cost_record->theoric_time * $component_cost_record->hour_cost * $tab_bom_cost_record->qty);
                        //pusho la riga testa -> non di acquisto
                        array_push($production_components, array(
                                'id'                => $component_cost_record->id_component_cost,
                                'code'              => $code,
                                'description'       => $description,
                                'qty'               => $tab_bom_cost_record->qty,
                                'weight'            => $component_cost_record->weight,
                                'material_cost'     => $component_cost_record->material_cost,
                                'effect_percentage' => $component_cost_record->effect_percentage,
                                'theoric_time'      => $component_cost_record->theoric_time,
                                'hour_cost'         => $component_cost_record->hour_cost,
                                'production_cost'   => $component_cost_record->production_cost,//$production_cost,
                                'component_cost'    => $component_cost_record->production_cost + $component_cost_record->material_cost,//$production_cost + $component_cost_record->material_cost,
                                'details'           => $details,
                            )
                        );

                        $totals['hour_cost'] += $component_cost_record->hour_cost;
                        $totals['production_cost'] += $component_cost_record->production_cost;
                        $totals['material_cost'] += $component_cost_record->material_cost;
                        $totals['component_cost'] += ($component_cost_record->production_cost + $component_cost_record->material_cost);
                    }
                }
            }
            //
        }

        //error_log(print_r($production_components, true));
        return array('production_components' => $production_components, 'totals' => $totals);
    }

    public function recursive_bom(&$details, $component_id, $component_table, $year){
        $subcomponents = $this->CI->utility->get_records($this->CI->db->tb_bom_cost, array('fk_parent_id' => $component_id, 'fk_parent_table' => $component_table));
        if(!$subcomponents){
            return;
        }

        foreach ($subcomponents as $subcomponent) {
            $primary_key = 'id_component';
            if($subcomponent->fk_child_table == 'tab_packaging'){
                $primary_key = 'id_packaging';
            }
            $subcomponent_anag = $this->CI->utility->get_record_by_id($primary_key, $subcomponent->fk_child_id, $subcomponent->fk_child_table);
            $subcomponent_cost = $this->CI->utility->get_records($this->CI->db->tb_component_cost, array('fk_item_id' => $subcomponent->fk_child_id, 'fk_item_table' => $subcomponent->fk_child_table, 'year' => $year));
            if($subcomponent_anag && $subcomponent_cost){
                $subcomponent_cost = $subcomponent_cost[0];

                $subdetails = array();
                array_push($subdetails,
                    array(
                        'code'              => $subcomponent_anag->code,
                        'description'       => $subcomponent_anag->ml_description,
                        'qty'               => $subcomponent->qty,
                        'weight'            => $subcomponent_cost->weight,
                        'cost'              => $subcomponent_cost->cost,
                        'theoric_time'      => $subcomponent_cost->theoric_time,
                        'effect_percentage' => $subcomponent_cost->effect_percentage,
                        'hour_cost'         => $subcomponent_cost->hour_cost,
                        'production_cost'   => $subcomponent_cost->production_cost,
                        'is_purchase'       => $subcomponent_cost->is_purchase,
                    )
                );

                $this->recursive_bom($subdetails[0], $subcomponent->fk_child_id, $subcomponent->fk_child_table, $year);
                $details['subdetails'][] = $subdetails[0];
                $details['subdetails'] = $this->sort_production_subcomponents_cost($details['subdetails']);
            }

        }
    }

    /**
     *
     */
    public function get_purchase_components_cost($id_article_cost, $year) {
        //tab_component_cost tipo "ACQ"
        //testa
        /*
        codice
        descrizione
        quantità di impiego			(da tab_bom_cost)
        peso					(da tab_component)
        costo unitario (€/kg)			tab_component_cost.purchase_price
        costo del componente			tab_component_cost.cost
        */

        $purchase_components = array();
        $totals = array(
            'cost' => 0,
            'component_cost' => 0,
        );

        //alla fine riga con somma del totale
        $article_cost_record = $this->CI->utility->get_record_by_id($this->primary_key, $id_article_cost, $this->class_tb);
        //$article_cost_detail_record = $this->utility->get_record_by_id('fk_article_cost_id', $id_article_cost, $this->CI->db->tb_article_cost_detail);

        //con questi due vado a leggere la component_cost e la component con il codice articolo
        $article_table = $article_cost_record->fk_item_table;
        $article_id = $article_cost_record->fk_item_id;

        //prendo da tab bom cost con articolo
        $tab_bom_cost_records = $this->CI->utility->get_records($this->CI->db->tb_bom_cost, array('fk_parent_table' => $article_table, 'fk_parent_id' => $article_id));

        /*mega query fatta solo per ordinare i codici, causa tabelle troppo sparse e senza chiavi per componenti e packaging*/
        $sql = "SELECT * 
                    FROM 
                        (SELECT tab_bom_cost.*, tab_component.code
                        FROM tab_bom_cost
                        JOIN tab_component ON fk_child_id = id_component
                        WHERE fk_parent_table = 'tab_article' AND fk_parent_id = {$article_id} AND fk_child_table = 'tab_component'
                        ORDER BY tab_component.code ASC ) AS tab_comps
                UNION ALL
                SELECT * 
                    FROM
                        (SELECT tab_bom_cost.*, tab_packaging.code
                        FROM tab_bom_cost
                        JOIN tab_packaging ON fk_child_id = id_packaging
                        WHERE fk_parent_table = 'tab_article' AND fk_parent_id = {$article_id} AND fk_child_table = 'tab_packaging'
                        ORDER BY tab_packaging.code ASC ) AS tab_packs
                ORDER BY code";
        $tab_bom_cost_records = $this->CI->db->query($sql)->result();

        if(!empty($tab_bom_cost_records)) {
            //
            foreach ($tab_bom_cost_records as $tab_bom_cost_record) {
                //ciclo gli elementi legati all'articolo
                $component_child_table = $tab_bom_cost_record->fk_child_table;
                $component_child_id = $tab_bom_cost_record->fk_child_id;
                //per ognuno che rientra in tab_component_cost
                if ($component_child_table == 'tab_component' || $component_child_table == 'tab_packaging') {
                    //vado a prendere il suo record in component cost
                    $component_cost_record = array();
                    $component_cost_records = $this->CI->utility->get_records($this->CI->db->tb_component_cost, array('fk_item_table' => $component_child_table, 'fk_item_id' => $component_child_id, 'year' => $year));
                    if(!empty($component_cost_records)) {
                        $component_cost_record = $component_cost_records[0];
                    }

                    if(!empty($component_cost_record)) {
                        //se è di acquisto allora lo prendo
                        if ($component_cost_record->is_purchase) {
                            $component = array();
                            $code = '';
                            $description = '';
                            if($component_child_table == 'tab_component') {
                                $component = $this->CI->utility->get_record_by_id('id_component', $component_child_id, $component_child_table);
                            } else if ($component_child_table == 'tab_packaging') {
                                $component = $this->CI->utility->get_record_by_id('id_packaging', $component_child_id, $component_child_table);
                            }
                            //
                            if(!empty($component)) {
                                $code = $component->code;
                                //parametrizzare in base alla lingua di sistema
                                $description = $component->it_ml_description;
                            }
                            //
                            array_push($purchase_components, array(
                                    'code' => $code,
                                    'description' => $description,
                                    'qty' => $tab_bom_cost_record->qty,
                                    'weight' => $component_cost_record->weight,
                                    'cost' => $component_cost_record->purchase_price,
                                    'component_cost' => $component_cost_record->cost
                                )
                            );

                            $totals['cost'] += $component_cost_record->purchase_price;
                            $totals['component_cost'] += $component_cost_record->cost;
                        }
                    }
                }
            }
        }
        //error_log(print_r($purchase_components, true));
        return array('purchase_components' => $purchase_components, 'totals' => $totals);
    }

    /**
     *
     */
    public function get_assembly_costs($id_article_cost, $year) {
        /*
        Si compone di una sola riga in cui riportare:
        la percentuale di carico 		(tab_assemby.assembly_perc)
        il tempo di assemblaggio 	(tab_article_assembly_time.theoric_time)
        il costo orario del reparto 	(da configurazione)
        Il costo dell’imballo		(tempo * costo orario)
        Verrà evidenziato il totale costo di assemblaggio.
        */

        $article_cost = $this->CI->utility->get_record_by_id('id_article_cost', $id_article_cost, $this->CI->db->tb_article_cost);

        $assembly_perc = $this->CI->utility->get_field_value('assembly_perc',$this->CI->db->tb_assembly, array('year' => $year));
        $assembly_time = $this->CI->utility->get_field_value('theoric_time',$this->CI->db->tb_article_assembly_time, array('fk_item_table' => 'tab_article', 'fk_item_id' => $article_cost->fk_item_id));

        $cost_config = array();
        $cost_configs = $this->CI->utility->get_records($this->CI->db->tb_tab_cost_config, array('year' => $year));
        if(!empty($cost_configs)) {
            $cost_config = $cost_configs[0];
        }
        if(!empty($cost_config)) {
            //$department_hour_cost = $cost_config->IC_department * ($cost_config->tax_perc / 100);
            $department_hour_cost = $cost_config->IC_department;
        } else {
            $department_hour_cost = 0;
        }

        $assembly_perc = !empty($assembly_perc) ? $assembly_perc : 0;
        $assembly_time = !empty($assembly_time) ? $assembly_time : 0;
        $department_hour_cost = !empty($department_hour_cost) ? $department_hour_cost : 0;
        $packaging_cost = ($assembly_time/60 * $department_hour_cost);

        return array(
            'charge_percentage' => $assembly_perc,
            'assembly_time' => $assembly_time,
            'department_hour_cost' => $department_hour_cost,
            'packaging_cost' => $packaging_cost,
            'total' => $packaging_cost + ($packaging_cost * ($assembly_perc / 100)),
        );
    }

    /**
     *
     */
    public function get_indirect_costs($year) {
        /*
        Si compone di tre righe riportanti:
        Fatturato anno precedente		tab_indirect_cost.indirect_cost_tot
        Costi indiretti anno precedente		tab_indirect_cost.revenue
        percentuale incidenza costi indiretti	tab_indirect_cost.indirect_perc

        La tabella tab_indirect_cost va interrogata per l’anno indicato nel filtro di accesso alla procedura.
        Il totale della sezione costi indiretti è dato da percentuale incidenza costi indiretti per il prezzo di vendita (tab_indirect_cost.indirect_perc * tab_article_cost_detail.sale_price).
         */

        $indirect_cost = $this->CI->utility->get_record_by_id('year', $year, $this->CI->db->tb_indirect_cost);
        $indirect_cost_tot = (!empty($indirect_cost) ? $indirect_cost->indirect_cost_tot : 0);
        $indirect_revenue = (!empty($indirect_cost) ? $indirect_cost->revenue : 0);
        $indirect_perc = (!empty($indirect_cost) ? $indirect_cost->indirect_perc : 0);

        return array(
            'indirect_cost_tot' => $indirect_cost_tot,
            'indirect_revenue' => $indirect_revenue,
            'indirect_perc' => $indirect_perc,
            'year' => $year,
            'total' => 0,
        );
    }

    /**
     *
     */
    public function get_total_costs($id_article_cost, $price_list_code) {
        /*
         Si compone di sei righe riportanti:
         totale costi ante tasse			Somma dei totali delle sezioni 1-2-3-4 = tab_article_cost_detail.cost
         prezzo di vendita			tab_article_cost_detail.sale_price
         percentuale di ricarico			(Prezzo di vendita - totale costi ante tasse) / prezzo di vendita %
         imposte				tab_article_cost_detail.taxes
         totale costi dopo tasse			tab_article_cost_detail.final_cost
         percentuale di ricarico dopo le tasse 	(Prezzo di vendita - totale costi) / prezzo di vendita %

         Evidenziare il costo finale.
         */
        $total_cost_no_taxes = 0;
        $sale_price = 0;
        $charge_percentage_no_taxes = 0;
        $taxes = 0;
        $total_cost = 0;
        $charge_percentage= 0;

        //$price_list_id = $this->CI->utility->get_field_value('id_price_list', $this->CI->db->tb_price_list, array('code' => $price_list_code));
        $price_list_id = $price_list_code;
        $article_cost = $this->CI->utility->get_record_by_id('id_article_cost', $id_article_cost, $this->CI->db->tb_article_cost);

        if(!empty($article_cost)) {
            $article_cost_detail = array();
            $article_cost_details = $this->CI->utility->get_records($this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $article_cost->id_article_cost, 'fk_price_list_id' => $price_list_id));
            if(!empty($article_cost_details)) {
                $article_cost_detail = $article_cost_details[0];
            }

            if(!empty($article_cost_detail)) {
                $total_cost_no_taxes = $article_cost_detail->cost;
                $sale_price = $article_cost_detail->sale_price;
                $charge_percentage_no_taxes = ($sale_price != 0 ? (($sale_price - $total_cost_no_taxes) / $sale_price) : 0);
                $taxes = $article_cost_detail->taxes;
                $total_cost = $article_cost_detail->final_cost;
                $charge_percentage= ($sale_price != 0 ? (($sale_price - $total_cost) / $sale_price) : 0);
            }
        }

        return array(
            'total_cost_no_taxes' => $total_cost_no_taxes,
            'sale_price' => $sale_price,
            'charge_percentage_no_taxes' => $charge_percentage_no_taxes,
            'taxes' => $taxes,
            'total_cost' => $total_cost,
            'charge_percentage' => $charge_percentage
        );
    }

    /**
     * @param $id_article_cost
     * @param $price_list_id
     * @return mixed
     */
    public function create_article_cost_v1($id_article_cost, $price_list_code) {
        //$price_list_id = $this->CI->utility->get_field_value('id_price_list', $this->CI->db->tb_price_list, array('code' => $price_list_code));
        $price_list_id = $price_list_code;
        if(!$price_list_id) {
            return false;
        }

        $article_cost = $this->CI->utility->get_record_by_id('id_article_cost', $id_article_cost, $this->CI->db->tb_article_cost);
        $article_cost_details = $this->CI->utility->get_records($this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $id_article_cost, 'fk_price_list_id' => $price_list_id));

        $c_article_cost_details = (!empty($article_cost_details) ? count($article_cost_details) : 0);
        $article_cost_detail_inserted = 0;

        $article_cost_v1_data = array(
            'fk_item_table' => $article_cost->fk_item_table,
            'fk_item_id' => $article_cost->fk_item_id,
            'year' => $article_cost->year,
            'version' => 1,
            'component_acq_cost' => $article_cost->component_acq_cost,
            'component_prd_cost' => $article_cost->component_prd_cost,
            'assembly_cost' => $article_cost->assembly_cost,
            //'currency' => 'EU',
            'user_insert' => ($this->CI->session->userdata('username')) ? $this->CI->session->userdata('username') : 'ERP',
            'date_insert' => date('Y-m-d H:i:s'),
        );

        $this->CI->db->insert($this->CI->db->tb_article_cost, $article_cost_v1_data);
        $article_cost_v1_id = $this->CI->db->insert_id();

        if($article_cost_v1_id) {
            for($i = 0; $i < $c_article_cost_details; $i++) {
                $article_cost_detail_v1_data = array(
                    'fk_article_cost_id' => $article_cost_v1_id,
                    'fk_price_list_id' => $article_cost_details[$i]->fk_price_list_id,
                    'sale_price' => $article_cost_details[$i]->sale_price,
                    'indirect_cost' => $article_cost_details[$i]->indirect_cost,
                    'cost' => $article_cost_details[$i]->cost,
                    'taxes' => $article_cost_details[$i]->taxes,
                    'final_cost' => $article_cost_details[$i]->final_cost,
                    'currency' => $article_cost_details[$i]->currency,
                    'user_insert' => ($this->CI->session->userdata('username')) ? $this->CI->session->userdata('username') : 'ERP',
                    'date_insert' => date('Y-m-d H:i:s'),
                );

                $this->CI->db->insert($this->CI->db->tb_article_cost_detail, $article_cost_detail_v1_data);
                $article_cost_detail_v1_id = $this->CI->db->insert_id();

                if($article_cost_detail_v1_id) {
                    $article_cost_detail_inserted++;
                }
            }
        }
        //se c'è l'id dell article_cost, se il conto dei suoi details è > 0 e il conto di quelli v1 inseriti è uguale al conto dei suoi dettagli allora restituisco l'id, altrimenti false
        return ($article_cost_v1_id && $c_article_cost_details > 0 && ($article_cost_detail_inserted == $c_article_cost_details)) ? $article_cost_v1_id : false;
    }

    /**
     * Update
     *
     * @return bool
     * */
    public function update($article_cost_v1_id, $price_list_code, $data_to_update) {
        //$price_list_id = $this->CI->utility->get_field_value('id_price_list', $this->CI->db->tb_price_list, array('code' => $price_list_code));
        $price_list_id = $price_list_code;
        if(!$price_list_id) {
            return false;
        }

        $sale_price_to_update = str_replace(',', '.', $data_to_update['sale_price']);
        $charge_percentage_no_taxes_to_update = str_replace(',', '.', $data_to_update['charge_percentage_no_taxes']);

        $article_cost_detail_v1 = array();
        $article_cost_details_v1 = $this->CI->utility->get_records($this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $article_cost_v1_id, 'fk_price_list_id' => $price_list_id));
        if(!empty($article_cost_details_v1)) {
            $article_cost_detail_v1 = $article_cost_details_v1[0];
        }

        if(!empty($article_cost_detail_v1)) {
            $sale_price = $article_cost_detail_v1->sale_price;
            $cost = $article_cost_detail_v1->cost;

            if($article_cost_detail_v1->sale_price != $sale_price_to_update) {
                //aggiorno sale price
                $this->CI->db->update($this->CI->db->tb_article_cost_detail, array('sale_price' => $sale_price_to_update), array('id_article_cost_detail' => $article_cost_detail_v1->id_article_cost_detail));
            }

            $charge_percentage_no_taxes = ($sale_price != 0 ? (($sale_price - $cost) / $sale_price) : 0);
            if($charge_percentage_no_taxes != $charge_percentage_no_taxes_to_update) { //è cambiato devo aggiornare il $article_cost_detail_v1->cost
                //se questo è cambiato allora devo risalire ad un nuovo sale price e ad un nuovo costo
                $cost_to_update = $sale_price_to_update - (($charge_percentage_no_taxes_to_update / 100) * $sale_price_to_update);
                $this->CI->db->update($this->CI->db->tb_article_cost_detail, array('cost' => $cost_to_update), array('id_article_cost_detail' => $article_cost_detail_v1->id_article_cost_detail));
            }

            return true;
        }
        return false;
    }

    /**
     *
     */
    public function print_pdf_cost_card($product_cost_card_data) {
        $this->CI->load->library('pdf');

        $pdf_filename = 'Costo prodotto [' . $product_cost_card_data['item_code'] . '] ' . $product_cost_card_data['item_description'];

        $pdf = new Pdf();
        $pdf->SetTitle($pdf_filename);
        $pdf->SetFont('dejavusans', '', 9, '', true);
        $pdf->SetMargins(10, 10, 10);
        $pdf->SetAutoPageBreak(TRUE, 10);
        $pdf->setPrintHeader(false);
        $pdf->setPrintFooter(false);
        // set image scale factor
        $pdf->setImageScale(PDF_IMAGE_SCALE_RATIO);

        $pdf->AddPage('P', 'A4');

        $pdf_html = $this->compose_html($product_cost_card_data);
        //

        $pdf->writeHTML($pdf_html, true, false, true, false, '');
        $pdf->lastPage();

        $pdf->Output($pdf_filename, 'I');
    }


    /**
     *
     */
    //CSS
    var $title_style            = 'font-size: 16px; text-align: center;';
    var $sim_title_style        = 'font-size: 10px; text-align: right;';
    var $table_label_style      = 'font-size: 9px;';
    var $small_label_style      = 'font-size: 7px;';
    var $table_border_style     = 'border: 0.5px solid #ccc;';
    var $table_cellpadding      = "4";
    var $td_number_style        = 'text-align: right;';
    var $td_border_style        = 'border: 1px solid #ccc;';
    var $text_center            = 'text-align: center;';
    var $text_right             = 'text-align: right;';
    var $text_mid               = 'vertical-align: middle!important;';
    var $code_col               = 'width: 11%!important;';
    var $descr_col              = 'width: 27%!important;';
    var $material_cost_col      = 'width: 25%!important;';
    var $production_cost_col    = 'width: 26%!important;';
    var $tot_col                = 'width: 11%!important;';
    var $qty_col                = 'width: 5%!important;';
    var $weight_col             = 'width: 6%!important;';
    var $unit_cost_col          = 'width: 7%!important;';
    var $prod_cost_col          = 'width: 7%!important;';
    var $thick_border_top       = 'border-top: solid 2px black!important;';
    var $thick_border_left      = 'border-left: solid 2px black!important;';
    var $thick_border_right     = 'border-right: solid 2px black!important;';
    var $thick_border_bottom    = 'border-bottom: solid 2px black!important;';
    var $bigger_text            = 'font-size: 12px!important;';
    var $bold                   = 'font-weight:600 !important;';
    var $lightblue_bg           = 'background-color: lightblue!important;';

    private function compose_html($product_cost_card_data) {
        $html = '<table style="' . $this->table_border_style . '" cellpadding="' . $this->table_cellpadding . '">';
        $html .=    '<tbody>';
        $html .=        '<tr>';
        $html .=            '<td style="' . $this->title_style . '">';
        $html .=                $product_cost_card_data['item_code'];
        $html .=            '</td>';
        $html .=            '<td colspan="2" style="' . $this->title_style . '">';
        $html .=                $product_cost_card_data['item_description'];
        $html .=            '</td>';
        $html .=            '<td style="' . $this->title_style . '">';
        $html .=                date('d/m/Y', strtotime(is_null($product_cost_card_data['date_modify']) ? $product_cost_card_data['date_insert'] : $product_cost_card_data['date_modify']));
        $html .=            '</td>';
        //$html .=            '<td style="' . $this->title_style . '">';
        //$html .=                '[' . $product_cost_card_data['item_code'] . '] - ' . $product_cost_card_data['item_description'];
        //$html .=            '</td>';

        /*SIMULAZIONE?*/
        if($product_cost_card_data['version'] > 0) {
            $html .=            '<td style="' . $this->sim_title_style . '">';
            $html .=                strtoupper(lang('simulation') . ' ' . strtoupper(lang('calculated')) . ' ' . strtoupper(lang('from')) . ': ' . strtoupper($product_cost_card_data['user_modify']) . ' ' . strtoupper(lang('on')) . ' ' . strtoupper($product_cost_card_data['date_modify']));
            $html .=            '</td>';
        }

        $html .=        '</tr>';
        $html .=    '</tbody>';
        $html .= '</table>';

        $html .= '<br><br>';

        /*COMPONENTI DI PRODUZIONE*/
        $prod_components_res = $this->print_prod_components($product_cost_card_data);
        $tot_prod_components = $prod_components_res['tot_prod_components'];
        $html .= $prod_components_res['html'];

        /*COSTI COMPONENTI DI ACQUISTO*/
        $purc_components_res = $this->print_acq_components($product_cost_card_data);
        $tot_purc_components = $purc_components_res['tot_purc_components'];
        $html .= $purc_components_res['html'];

        /*COSTI ASSEMBLAGGIO*/
        $asse_components_res = $this->print_assembly_cost($product_cost_card_data);
        $tot_asse_components = $asse_components_res['tot_asse_components'];
        $html .= $asse_components_res['html'];

        /*COSTI INDIRETTI*/
        $indi_components_res = $this->print_indirect_cost($product_cost_card_data);
        $tot_indi_components = $indi_components_res['tot_indi_components'];
        $html .= $indi_components_res['html'];

        /*TOTALE*/
        $tot_no_taxes       = $tot_prod_components + $tot_purc_components + $tot_asse_components;
        $sale_price         = $this->CI->utility->get_field_value('sale_price', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $sale_price_orig    = $this->CI->utility->get_field_value('sale_price_orig', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $taxes              = $this->CI->utility->get_field_value('taxes', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $year               = $this->CI->utility->get_field_value('year', $this->CI->db->tb_article_cost, array('id_article_cost' => $product_cost_card_data['id_article_cost']));
        $cost_config_perc   = $this->CI->utility->get_field_value('tax_perc', $this->CI->db->tb_tab_cost_config, array('year' => $year));
        $taxes              = ($sale_price - ($tot_no_taxes + $tot_indi_components)) * $cost_config_perc/100;
        if($taxes < 0){
            $taxes = 0;
        }
        $html .= $this->print_total_cost($product_cost_card_data, $tot_no_taxes, $tot_indi_components, $taxes, $sale_price, $sale_price_orig);

        return $html;
    }

    public function print_prod_components($product_cost_card_data){
        $html = '<table style="' . $this->table_label_style . '" cellpadding="2">';
        $html .= '<thead>';
        $html .= $this->print_prod_components_header();
        $html .= '</thead>';
        $html .= '<tbody>';

        $tot_prod_components = 0;
        if(!empty($product_cost_card_data['production_components_cost'])) {
            if(!empty($product_cost_card_data['production_components_cost']['production_components'])) {
                //foreach ($product_cost_card_data['production_components_cost']['production_components'] as $production_component) {
                $counter = count($product_cost_card_data['production_components_cost']['production_components']);
                for ($i = 0; $i < $counter; $i++) {
                    if(substr($product_cost_card_data['production_components_cost']['production_components'][$i]['code'], 0, 3) != '60.') {
                        $tot_prod_components += $product_cost_card_data['production_components_cost']['production_components'][$i]['component_cost'] * $product_cost_card_data['production_components_cost']['production_components'][$i]['qty'];
                        $html .= $this->print_prod_components_main($product_cost_card_data['production_components_cost']['production_components'][$i]);

                        if(!empty($product_cost_card_data['production_components_cost']['production_components'][$i]['details'])) {
                            $detail_counter = count($product_cost_card_data['production_components_cost']['production_components'][$i]['details']);
                            for($j = 0; $j < $detail_counter; $j++) {
                                $detail = $product_cost_card_data['production_components_cost']['production_components'][$i]['details'][$j];
                                $html .= $this->print_prod_components_submain($detail);

                                $html .= $this->print_prod_components_subs($detail);
                            }
                        }
                    }

                }
            }
            //riga totale
            if(!empty($product_cost_card_data['production_components_cost']['totals'])) {
                $html .=   '<tr>
                                <td style="border-top: 0.5px solid black; width: 89%!important;"></td>
                                <td bgcolor="#eaeaea" style="font-size:8pt; width: 11%!important;text-align: right;border: 0.5px solid black;"><b>' . number_format($tot_prod_components, 2, ',', '.') . '</b></td>
                            </tr>';
            }
        }
        $html .= '</tbody>';
        $html .= '</table>';
        $html .= '<br><br>';

        $result['tot_prod_components']  = $tot_prod_components;
        $result['html']                 = $html;

        return $result;
    }

    public function print_prod_components_header(){
        $html = '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:8pt;width: 15%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('direct_costs'))    . '</strong></th>
                    <th style="font-size:8pt;width: 22%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('productions'))     . '</strong></th>
                    <th style="font-size:8pt;width: 25%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('materials'))       . '</strong></th>
                    <th style="font-size:8pt;width: 27%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('times'))           . '</strong></th>
                    <th style="font-size:8pt;width: 11%!important; text-align: center; border-left: 0.5px solid black; border-top: 0.5px solid black; border-right: 0.5px solid black"><strong>' . strtoupper(lang('total'))           . '</strong></th>
                </tr>';
        $html .= '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:6pt; width: 15%!important;text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 22%!important;text-align: center;border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 5%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('pz_short'))        . '</th>
                    <th style="font-size:6pt; width: 6%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('weight'))          . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('kg_cost'))         . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;border-right: 0.5px solid black;">' . strtoupper(lang('total'))           . '</th>
                    <th style="font-size:6pt; width: 6%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('label_time'))      . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('%_ineff_short'))   . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('h_cost_shorter'))  . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;border-right: 0.5px solid black;">' . strtoupper(lang('total'))           . '</th>
                    <th style="width: 11%!important;border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                 </tr>';

        return $html;
    }



    public function print_prod_components_main($elem){
        $tot_row_span       = 1 + count($elem['details']);
        $subtot_row_span    = 1;
        $subtot_acq_comp    = 0;
        foreach ($elem['details'] as $detail) {
            if(!empty($detail['subdetails'])){
                $tot_row_span += count($detail['subdetails']);
                foreach ($detail['subdetails'] as $subdetail) {
                    if(!$subdetail['is_purchase']){
                        $tot_row_span++;
                    }
                }
            }
            if($detail['is_purchase']){
                $subtot_row_span++;
                $subtot_acq_comp += $detail['cost'] *$detail['qty'] * $detail['weight'];
            }
        }

        $html = '<tr style="font-size: 6.5pt;border-top: 0.5px solid black">
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;">' . $elem['code']                           . '</td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 22%!important;"><strong>' . substr($elem['description'], 0, 30)                    . '</strong></td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 1px solid #ccc;  width: 5%!important;text-align: right;">' . number_format($elem['qty'], 0)  .'</td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 20%!important;"></td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 1px solid #ccc;  width: 6%!important;text-align: right;">' . $this->CI->utility->decimal_to_hour_min_sec($elem['theoric_time'])       . '</td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 1px solid #ccc;  width: 7%!important;text-align: right;">' . number_format($elem['effect_percentage'], 2, ',', '.')  . ' %</td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 1px solid #ccc;  width: 7%!important;text-align: right;">' . number_format($elem['hour_cost'], 2, ',', '.')          . '</td>
                    <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 7%!important;text-align: right;">' . number_format($elem['production_cost'], 2, ',', '.')    . '</td>';
        if($tot_row_span > $subtot_row_span){
            $html .=  '<td style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 5%!important; text-align: right;" rowspan="' . $subtot_row_span . '">' . number_format(($elem['production_cost'] + $subtot_acq_comp) * $elem['qty'], 2, ',', '.') . '</td>';
            $html .=  '<td style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 6%!important; text-align: right;" rowspan="' . $subtot_row_span . '">' . number_format($elem['component_cost'] * $elem['qty'], 2, ',', '.') . '</td>';
        }
        else{
            $html .=  '<td style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; width: 11%!important; text-align: right;">' . number_format($elem['component_cost'] * $elem['qty'], 2, ',', '.') . '</td>';
        }
        $html .= '</tr>';

        return $html;
    }

    public function print_prod_components_submain($detail){
        $subdetail_row_span     = 1;
        $subtot_acq_subdetails  = 0;
        if(!empty($detail['subdetails'])){
            foreach ($detail['subdetails'] as $subdetail) {
                if($subdetail['is_purchase']){
                    $subdetail_row_span++;
                    $subtot_acq_subdetails += $subdetail['cost'] * $subdetail['qty'] /** $subdetail['weight']*/;
                }
            }
        }

        $html = '<tr style="font-size: 6.5pt;">';
        if($detail['is_purchase']){
            $html .=   '<td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;">' . $detail['code'] . '</td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 22%!important;">' . substr($detail['description'], 0, 30) . '</td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 5%!important;"></td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 6%!important;text-align: right;">' . number_format($detail['qty'], 4) . '</td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($detail['cost'], 4) . '</td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 7%!important;text-align: right;">' . number_format($detail['cost'] *$detail['qty'] /** $detail['weight']*/, 2, ',', '.') . '</td>
                        <td style="border-top: 1px solid #ccc; width: 6%!important;"></td>
                        <td style="border-top: 1px solid #ccc; width: 7%!important;"></td>
                        <td style="border-top: 1px solid #ccc; width: 7%!important;"></td>
                        <td style="border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 7%!important;"></td>
                        <td style="border-right: 0.5px solid black; width: 11%!important;"></td>';
        }
        else{
            $html .=   '<td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"><strong>' . $detail['code'] . '</strong></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 22%!important;"><strong>' . substr($detail['description'], 0, 30) . '</strong></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 5%!important;text-align: right;">' . number_format($detail['qty'], 0) . '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; width: 6%;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; width: 7%;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 7%;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 6%!important;text-align: right;">' . $this->CI->utility->decimal_to_hour_min_sec($detail['theoric_time']) . '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($detail['effect_percentage'], 2, ',', '.') . ' %</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($detail['hour_cost'], 2, ',', '.') . '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 7%!important;text-align: right;">' . number_format($detail['production_cost'], 2, ',', '.') . '</td>
                        <td style="font-size:6pt;border-right: 0.5px solid black; width: 5%!important;text-align: right;border-top: 1px solid darkgrey;" rowspan="' . $subdetail_row_span . '">' . number_format(($detail['production_cost'] + $subtot_acq_subdetails) * $detail['qty'], 2, ',', '.') . '</td>
                        <td style="border-right: 0.5px solid black; width: 6%!important;" rowspan="' . $subdetail_row_span . '"></td>';
        }
        $html .= '</tr>';

        return $html;
    }

    public function print_prod_components_subs($detail){
        if (array_key_exists('subdetails', $detail) && !empty($detail['subdetails'])) {
            $html = '';
            $subdetail_counter = count($detail['subdetails']);
            for($k = 0; $k < $subdetail_counter; $k++){
                $subdetail  = $detail['subdetails'][$k];
                $row_span   = 1 + ((array_key_exists('subdetails', $subdetail)) ? count($subdetail['subdetails']) : 0);
                $html .= '<tr style="font-size: 6.5pt;">';
                if($subdetail['is_purchase']){
                    $html .=   '<td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;">' . $subdetail['code'] . '</td>
                                <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 22%!important;">' . substr($subdetail['description'], 0, 30) . '</td>
                                <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 5%!important;"></td>
                                <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 6%!important;text-align: right;">' . number_format($subdetail['qty'], 4) . '</td>
                                <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($subdetail['cost'], 4, ',', '.') . '</td>
                                <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 7%!important;text-align: right;">' . number_format($subdetail['cost'] * $subdetail['qty'] /** $subdetail['weight']*/, 2, ',', '.') . '</td>
                                <td style="border-top: 1px solid #ccc; width: 6%!important;"></td>
                                <td style="border-top: 1px solid #ccc; width: 7%!important;"></td>
                                <td style="border-top: 1px solid #ccc; width: 7%!important;"></td>
                                <td style="border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 7%!important;"></td>
                                <td style="border-right: 0.5px solid black; width: 11%!important;"></td>';
                }
                else {
                    $html .=   '<td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"><strong>' . $subdetail['code'] . '</strong></td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 22%!important;"><strong>' . substr($subdetail['description'], 0, 30) . '</strong></td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 5%!important;text-align: right;">' . number_format($subdetail['qty'], 0) . '</td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; width: 6%;"></td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; width: 7%;"></td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 7%;"></td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 6%!important;text-align: right;">' . $this->CI->utility->decimal_to_hour_min_sec($subdetail['theoric_time']) . '</td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($subdetail['effect_percentage'], 2, ',', '.') . ' %</td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 1px solid #ccc; width: 7%!important;text-align: right;">' . number_format($subdetail['hour_cost'], 2, ',', '.') . '</td>
                                <td bgcolor="#eaeaea" style="font-size:6pt;font-style:italic; border-top: 1px solid darkgrey; border-right: 0.5px solid black; width: 7%!important;text-align: right;">' . number_format($subdetail['production_cost'], 2, ',', '.') . '</td>
                                <td style="font-size:6pt;border-right: 0.5px solid black; width: 5%!important;text-align: right;border-top: 1px solid darkgrey;" rowspan="' . $row_span . '">' . number_format($subdetail['cost'] * $subdetail['qty'], 2, ',', '.') . '</td>
                                <td style="border-right: 0.5px solid black; width: 6%!important;" rowspan="' . $row_span . '"></td>';
                }
                $html .= '</tr>';
                $html .= $this->print_prod_components_subs($subdetail);
            }
            return $html;
        }
    }

    public function print_acq_components($product_cost_card_data){
        $res = $this->print_acq_components_body($product_cost_card_data);
        $html = '<table style="' . $this->table_label_style . '" cellpadding="' . $this->table_cellpadding . '">';
        $html .= '<thead>';
        $html .= $this->print_acq_components_header();
        $html .= '</thead>';
        $html .= '<tbody>';
        $html .= $res['html'];
        $html .= '</tbody>';
        $html .= '</table>';
        $html .= '<br><br>';

        $result['tot_purc_components']  = $res['tot_purc_components'];
        $result['html']                 = $html;

        return $result;
    }

    public function print_acq_components_header(){
        $html = '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:8pt; width: 15%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('direct_costs'))    . '</strong></th>
                    <th style="font-size:8pt; width: 22%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('purchases'))     . '</strong></th>
                    <th style="font-size:8pt; width: 25%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('materials'))       . '</strong></th>
                    <th style="font-size:8pt; width: 27%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"></th>
                    <th style="font-size:8pt; width: 11%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black; border-right: 0.5px solid black"><strong>' . strtoupper(lang('total'))           . '</strong></th>
                </tr>';
        $html .= '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:6pt; width: 15%!important;text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 22%!important;text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 5%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('common_qty'))        . '</th>
                    <th style="font-size:6pt; width: 13%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('unit_cost_short'))         . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center; border-left: 1px solid #ccc;border-top: 1px solid #ccc;border-right: 0.5px solid black;">' . strtoupper(lang('total'))           . '</th>
                    <th style="font-size:6pt; width: 27%!important; border-right: 0.5px solid black; "></th>
                    <th style="font-size:6pt; width: 11%!important; border-right: 0.5px solid black; border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                 </tr>';
        return $html;
    }

    public function print_acq_components_body($product_cost_card_data){
        $html = '';
        $mid_categories_acq = array();
        if(!empty($product_cost_card_data['purchase_components_cost'])) {
            if (!empty($product_cost_card_data['purchase_components_cost']['purchase_components'])) {
                foreach ($product_cost_card_data['purchase_components_cost']['purchase_components'] as $purchase_components) {
                    $current_mid_category = explode('.', $purchase_components['code'])[1];
                    if(!array_key_exists($current_mid_category, $mid_categories_acq)){
                        $mid_categories_acq[$current_mid_category]['tot_rows']   = 0;
                        $mid_categories_acq[$current_mid_category]['sub_tot']    = 0;
                    }
                    $mid_categories_acq[$current_mid_category]['tot_rows']++;
                    $mid_categories_acq[$current_mid_category]['sub_tot'] += $purchase_components['qty'] * $purchase_components['cost'];
                }
            }
        }

        $tot_purc_components = 0;
        if(!empty($product_cost_card_data['purchase_components_cost'])) {
            if(!empty($product_cost_card_data['purchase_components_cost']['purchase_components'])) {
                $counter = 0;
                $old_mid_category = false;
                foreach ($product_cost_card_data['purchase_components_cost']['purchase_components'] as $purchase_components) {
                    $bg_color   = '';
                    $border_top = 'border-top: 1px solid #ccc;';
                    if($counter%2 != 0){
                        $bg_color   = 'bgcolor="#eaeaea"';
                    }
                    $current_mid_category = explode('.', $purchase_components['code'])[1];
                    if($old_mid_category != $current_mid_category){
                        $border_top = 'border-top: 0.5px solid black;';
                    }
                    $html .= '<tr>
                                <td ' . $bg_color . ' style="font-size:6pt; width: 15%!important;border-left: 0.5px solid black;border-right: 0.5px solid black;' . $border_top . '">' . $purchase_components['code'] . '</td>
                                <td ' . $bg_color . ' style="font-size:6pt; width: 22%!important;border-right: 0.5px solid black;' . $border_top . '">' . substr($purchase_components['description'], 0, 30) . '</td>
                                <td ' . $bg_color . ' style="font-size:6pt; width: 5%!important;text-align:right;border-right: 1px solid #ccc;' . $border_top . '">' . number_format($purchase_components['qty'], 2, ',', '.') . '</td>
                                <td ' . $bg_color . ' style="font-size:6pt; width: 13%!important;text-align:right;border-right: 1px solid #ccc;' . $border_top . '">' . number_format($purchase_components['cost'], 4, ',', '.'). '</td>
                                <td ' . $bg_color . ' style="font-size:6pt; width: 7%!important;text-align:right;border-right: 0.5px solid black;' . $border_top . '">' . number_format($purchase_components['qty'] * $purchase_components['cost'], 2, ',', '.') . '</td>';
                    if($old_mid_category != $current_mid_category){
                        $html .= '<td style="font-size:6pt; width: 27%!important;' . $border_top . '" rowspan="' . $mid_categories_acq[$current_mid_category]['tot_rows'] . '"></td>';
                        $html .= '<td style="font-size:6pt; width: 11%!important;text-align:right;border-left: 0.5px solid black;border-right: 0.5px solid black;' . $border_top . '" rowspan="' . $mid_categories_acq[$current_mid_category]['tot_rows'] . '">' . number_format($mid_categories_acq[$current_mid_category]['sub_tot'], 2, ',', '.') . '</td>';
                        $old_mid_category = $current_mid_category;
                    }
                    $html .= '</tr>';
                    $tot_purc_components += $purchase_components['qty'] * $purchase_components['cost'];
                    $counter++;
                }
            }
            //riga totale
            if(!empty($product_cost_card_data['purchase_components_cost']['totals'])) {
                $html .= '<tr>
                            <td style="border-top: 0.5px solid black; width: 89%!important;"></td>
                            <td bgcolor="#eaeaea" style="font-size:8pt; width: 11%!important;text-align: right;border: 0.5px solid black;"><b>' . number_format($tot_purc_components, 2, ',', '.') . '</b></td>
                        </tr>';
            }
        }

        $result['tot_purc_components']  = $tot_purc_components;
        $result['html']                 = $html;

        return $result;
    }

    public function print_assembly_cost($product_cost_card_data){
        $res = $this->print_assembly_cost_body($product_cost_card_data);
        $html = '<table style="' . $this->table_label_style . '" cellpadding="2">';
        $html .= '<thead>';
        $html .= $this->print_assembly_cost_header();
        $html .= '</thead>';
        $html .= '<tbody>';
        $html .= $res['html'];
        $html .= '</tbody>';
        $html .= '</table>';
        $html .= '<br><br>';

        $result['tot_asse_components']  = $res['tot_asse_components'];
        $result['html']                 = $html;

        return $result;
    }

    public function print_assembly_cost_header(){
        $html = '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:8pt; width: 15%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('direct_costs'))    . '</strong></th>
                    <th style="font-size:8pt; width: 22%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('assembly'))     . '</strong></th>
                    <th style="font-size:8pt; width: 25%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"></th>
                    <th style="font-size:8pt; width: 27%!important; text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('times'))           . '</strong></th>
                    <th style="font-size:8pt; width: 11%!important; text-align: center; border-left: 0.5px solid black; border-top: 0.5px solid black; border-right: 0.5px solid black"><strong>' . strtoupper(lang('total'))           . '</strong></th>
                </tr>';
        $html .= '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:6pt; width: 15%!important;text-align: center; border-right: 0.5px solid black; border-left: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 22%!important;text-align: center; border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 25%!important;border-right: 0.5px solid black;"></th>
                    <th style="font-size:6pt; width: 6%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('label_time'))      . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('%_carico_scarico_short'))   . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;">' . strtoupper(lang('%_ineff_short'))  . '</th>
                    <th style="font-size:6pt; width: 7%!important;text-align: center;border-left: 1px solid #ccc;border-top: 1px solid #ccc;border-right: 0.5px solid black;">' . strtoupper(lang('unit_cost_short'))           . '</th>
                    <th style="font-size:6pt; width: 11%!important;border-left: 0.5px solid black;border-right: 0.5px solid black;"></th>
                 </tr>';

        return $html;
    }

    public function print_assembly_cost_body($product_cost_card_data){
        $tot_assembly_cost = 0;
        $html = '';
        if(!empty($product_cost_card_data['assembly_costs'])) {
            $tot_assembly_cost = $product_cost_card_data['assembly_costs']['total'];
            $html .= '<tr>
                        <td bgcolor="#eaeaea" style="font-size:6pt;border-top: 0.5px solid black; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 22%!important; border-right: 0.5px solid black; border-top: 0.5px solid black">' . lang('assembly_detail'). '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 25%!important; border-right: 0.5px solid black; border-top: 0.5px solid black"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 6%!important;text-align: right;border-right: 1px solid #ccc;border-top: 0.5px solid black;">' . $this->CI->utility->decimal_to_hour_min_sec($product_cost_card_data['assembly_costs']['assembly_time']) . '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 7%!important;text-align: right;border-right: 1px solid #ccc;border-top: 0.5px solid black;">' . number_format($product_cost_card_data['assembly_costs']['charge_percentage'], 2, ',', '.') . '%</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 7%!important;text-align: right;border-right: 1px solid #ccc;border-top: 0.5px solid black;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 7%!important;text-align: right;border-right: 0.5px solid black;border-top: 0.5px solid black;">' . number_format($product_cost_card_data['assembly_costs']['department_hour_cost'], 2, ',', '.') . '</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt;width: 11%!important;text-align: right;border-right: 0.5px solid black;border-top: 0.5px solid black;">' . number_format($product_cost_card_data['assembly_costs']['total'], 2, ',', '.') . '</td>
                    </tr>';
        }
        /*KITs*/
        $tot_kits_cost      = 0;
        $assembly_dep_perc  = 0;
        $year               = $this->CI->utility->get_field_value('year', $this->CI->db->tb_article_cost, array('id_article_cost' => $product_cost_card_data['id_article_cost']));
        $department_id      = $this->CI->utility->get_field_value('id_department', $this->CI->db->tb_department, array('code' => 'CFZ'));
        if($year && $department_id) {
            $effect_percentage = $this->CI->utility->get_field_value('effect_percentage', $this->CI->db->tb_tab_department_worked, array('year' => $year, 'fk_department_id' => $department_id));
            if ($effect_percentage) {
                $assembly_dep_perc = $effect_percentage;
            }
        }
        if(!empty($product_cost_card_data['production_components_cost'])) {
            if (!empty($product_cost_card_data['production_components_cost']['production_components'])) {
                $counter = 0;
                foreach ($product_cost_card_data['production_components_cost']['production_components'] as $production_component) {
                    if (substr($production_component['code'], 0, 3) == '60.') {
                        $bg_color = '';
                        if($counter%2 != 0){
                            $bg_color = 'bgcolor="#eaeaea"';
                        }
                        $html .= '<tr ' . $bg_color . '>
                                    <td style="font-size:6pt;border-top: 1px solid #ccc; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;">' . $production_component['code'] . '</td>
                                    <td style="font-size:6pt;width: 22%!important; border-right: 0.5px solid black; border-top: 1px solid #ccc;">' . substr($production_component['description'], 0, 30) . '</td>
                                    <td style="font-size:6pt;width: 25%!important; border-right: 0.5px solid black; border-top: 1px solid #ccc;"></td>
                                    <td style="font-size:6pt;width: 6%!important;text-align: right;border-right: 1px solid #ccc;border-top: 1px solid #ccc;">' . $this->CI->utility->decimal_to_hour_min_sec($production_component['theoric_time']) . '</td>
                                    <td style="font-size:6pt;width: 7%!important;text-align: right;border-right: 1px solid #ccc;border-top: 1px solid #ccc;"></td>
                                    <td style="font-size:6pt;width: 7%!important;text-align: right;border-right: 1px solid #ccc;border-top: 1px solid #ccc;">' . number_format($assembly_dep_perc, 2, ',', '.') . '%</td>
                                    <td style="font-size:6pt;width: 7%!important;text-align: right;border-right: 0.5px solid black;border-top: 1px solid #ccc;">' . number_format($production_component['hour_cost'], 2, ',', '.') . '</td>
                                    <td style="font-size:6pt;width: 11%!important;text-align: right;border-right: 0.5px solid black;border-top: 1px solid #ccc;">' . number_format($production_component['production_cost'], 2, ',', '.') . '</td>
                                </tr>';
                        $tot_kits_cost += $production_component['production_cost'] * $production_component['qty'];
                        $counter++;
                    }
                }
            }
        }
        //riga totale
        if(isset($product_cost_card_data['assembly_costs']['total'])) {
            $html .= '<tr>
                        <td style="border-top: 0.5px solid black; width: 89%!important;"></td>
                        <td bgcolor="#eaeaea" style="font-size:8pt; width: 11%!important;text-align: right;border: 0.5px solid black;"><b>' . number_format($tot_kits_cost + $product_cost_card_data['assembly_costs']['total'], 2, ',', '.') . '</b></td>
                    </tr>';
        }

        $result['tot_asse_components']  = $tot_kits_cost + $product_cost_card_data['assembly_costs']['total'];
        $result['html']                 = $html;

        return $result;
    }

    public function print_indirect_cost($product_cost_card_data){
        $res = $this->print_indirect_cost_body($product_cost_card_data);
        $html = '<table style="' . $this->table_label_style . '" cellpadding="' . $this->table_cellpadding . '">';
        $html .= '<thead>';
        $html .= $this->print_indirect_cost_header();
        $html .= '</thead>';
        $html .= '<tbody>';
        $html .= $res['html'];
        $html .= '</tbody>';
        $html .= '</table>';
        $html .= '<br><br>';

        $result['tot_indi_components']  = $res['tot_indi_components'];
        $result['html']                 = $html;

        return $result;
    }

    public function print_indirect_cost_header(){
        $html = '<tr bgcolor="#1b8fd3" style="color: white;">
                    <th style="font-size:8pt; width: 15%!important; text-align: center; border-right: 0.5px solid black;border-left: 0.5px solid black; border-top: 0.5px solid black"><strong>' . strtoupper(lang('indirect_cost'))    . '</strong></th>
                    <th style="font-size:8pt; width: 22%!important; text-align: center; border-right: 0.5px solid black;border-left: 0.5px solid black; border-top: 0.5px solid black"></th>
                    <th style="font-size:8pt; width: 52%!important; text-align: center; border-right: 0.5px solid black;border-left: 0.5px solid black; border-top: 0.5px solid black"></th>
                    <th style="font-size:8pt; width: 11%!important; text-align: center; border-left: 0.5px solid black; border-top: 0.5px solid black; border-right: 0.5px solid black"><strong>' . strtoupper(lang('total'))           . '</strong></th>
                </tr>';

        return $html;
    }

    public function print_indirect_cost_body($product_cost_card_data){
        if(!empty($product_cost_card_data['indirect_costs'])) {

            $html = '<tr>
                        <td style="font-size:6pt; border-top: 0.5px solid black; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"></td>
                        <td style="font-size:6pt; border-top: 0.5px solid black; border-right: 0.5px solid black; width: 22%!important;">' . strtoupper(lang('revenue')) . ' ' . $product_cost_card_data['indirect_costs']['year'] . '</td>
                        <td style="font-size:6pt; width: 52%!important; text-align: center; border-left: 0.5px solid black; border-top: 0.5px solid black"></td>
                        <td style="font-size:6pt; width: 11%!important; text-align: right; border-left: 0.5px solid black; border-top: 0.5px solid black; border-right: 0.5px solid black">' . number_format($product_cost_card_data['indirect_costs']['indirect_revenue'], 2, ',', '.') . '</td>
                    </tr>';
            $html .= '<tr>
                        <td bgcolor="#eaeaea" style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 22%!important;">' . strtoupper(lang('indirect_cost_detail')) . ' ' . $product_cost_card_data['indirect_costs']['year'] .'</td>
                        <td bgcolor="#eaeaea" style="font-size:6pt; width: 52%!important; text-align: center; border-top: 1px solid #ccc"></td>
                        <td bgcolor="#eaeaea" style="font-size:6pt; width: 11%!important; text-align: right; border-left: 0.5px solid black; border-top: 1px solid #ccc; border-right: 0.5px solid black">' . number_format($product_cost_card_data['indirect_costs']['indirect_cost_tot'], 2, ',', '.') . '</td>
                    </tr>';
            $html .= '<tr>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; border-left: 0.5px solid black; width: 15%!important;"></td>
                        <td style="font-size:6pt; border-top: 1px solid #ccc; border-right: 0.5px solid black; width: 22%!important;">' . strtoupper(lang('effect_long_descr')) . '</td>
                        <td style="font-size:6pt; width: 52%!important; text-align: center; border-left: 0.5px solid black; border-top: 1px solid #ccc"></td>
                        <td style="font-size:6pt; width: 11%!important; text-align: right; border-left: 0.5px solid black; border-top: 1px solid #ccc; border-right: 0.5px solid black">' . number_format($product_cost_card_data['indirect_costs']['indirect_perc'], 2, ',', '.') . '</td>
                    </tr>';
        }

        if(isset($product_cost_card_data['indirect_costs']['total'])) {
            $sale_price         = $this->CI->utility->get_field_value('sale_price', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
            $tot_indirect_cost  = $sale_price * $product_cost_card_data['indirect_costs']['indirect_perc']/100;
            $html .= '<tr>
                        <td style="border-top: 0.5px solid black; width: 89%!important;"></td>
                        <td bgcolor="#eaeaea" style="font-size:8pt; width: 11%!important;text-align: right;border: 0.5px solid black;"><b>' . number_format($tot_indirect_cost, 2, ',', '.') . '</b></td>
                    </tr>';
        }

        $result['tot_indi_components']  = $tot_indirect_cost;
        $result['html']                 = $html;

        return $result;
    }

    public function print_total_cost($product_cost_card_data, $tot_no_taxes, $tot_indirect_cost, $taxes, $sale_price, $sale_price_orig){
        $html = '<table style="' . $this->table_label_style . '" cellpadding="' . $this->table_cellpadding . '">
                    <thead>
                        <tr>
                            <th bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></th>
                            <th bgcolor="#1b8fd3" style="font-size:8pt; width: 40%; text-align: center; color: white; border: 0.5px solid black"><strong>' . strtoupper(lang('riepilogo')) . '</strong></th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td bgcolor="#eaeaea" style="font-size:8pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc;">' . strtoupper(lang('direct_costs')) . '</td>
                            <td bgcolor="#eaeaea" style="font-size:8pt; width: 18%;text-align: right; border-right: 0.5px solid black;">' . number_format($tot_no_taxes, 2, ',', '.') . '</td>
                        </tr>
                        <tr>
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td style="font-size:8pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-top: 1px solid #ccc;">' . strtoupper(lang('indirect_cost')) . '</td>
                            <td style="font-size:8pt; width: 18%;text-align: right; border-right: 0.5px solid black; border-top: 1px solid #ccc;">' . number_format($tot_indirect_cost, 2, ',', '.'). '</td>
                        </tr>
                        <tr>
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td bgcolor="#eaeaea" style="font-size:6pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-top: 1px solid #ccc;">' . strtoupper(lang('taxes')) . '</td>
                            <td bgcolor="#eaeaea" style="font-size:6pt; width: 18%;text-align: right; border-right: 0.5px solid black; border-top: 1px solid #ccc;">' . number_format($taxes, 2, ',', '.'). '</td>
                        </tr>';
        if(isset($product_cost_card_data['total_costs']['total_cost'])) {
            $tot_post_taxes     = $tot_no_taxes + $taxes + $tot_indirect_cost;
            $perc_post_taxes    = 0;
            if($tot_post_taxes > 0){
                $perc_post_taxes    = ($sale_price - $tot_post_taxes) / $tot_post_taxes * 100;
            }
            $html .=    '<tr style="background-color: white;">
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td style="font-size:8pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-top: 0.5px solid black;"><strong>' . strtoupper(lang('total_taxes')) . '</strong></td>
                            <td style="font-size:10pt; width: 18%;text-align: right; border-right: 0.5px solid black; border-top: 0.5px solid black;"><strong>' . number_format($tot_post_taxes, 2, ',', '.') . '</strong></td>
                        </tr>
                        <tr>
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td bgcolor="#eaeaea" style="font-size:8pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-top: 1px solid #ccc;"><strong>' . strtoupper(lang('sale_price')) . '</strong></td>
                            <td bgcolor="#eaeaea" style="font-size:10pt; width: 18%;text-align: right; border-right: 0.5px solid black; border-top: 1px solid #ccc;"><strong>' . number_format($sale_price, 2, ',', '.') . '</strong></td>
                        </tr>
                        <tr style="color: red;">
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td style="font-size:8pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-top: 1px solid #ccc; border-bottom: 0.5px solid black;"><strong>' . strtoupper(lang('charge_percentage_taxes')) . '</strong></td>
                            <td style="font-size:9pt; width: 9%;text-align: right; border-right: 1px solid #ccc; border-top: 1px solid #ccc; border-bottom: 0.5px solid black;"><strong>€' . number_format($sale_price - $tot_post_taxes, 2, ',', '.') . '</strong></td>
                            <td style="font-size:9pt; width: 9%;text-align: right; border-right: 0.5px solid black; border-top: 1px solid #ccc; border-bottom: 0.5px solid black;"><strong>' . number_format($perc_post_taxes, 2, ',', '.') . '%</strong></td>
                        </tr>
                        <tr>
                            <td bgcolor="white" style="width: 60%!important; border-right: 0.5px solid black;"></td>
                            <td bgcolor="#eaeaea" style="font-size:6pt; width: 22%; border-left: 0.5px solid black; border-right: 1px solid #ccc; border-bottom: 0.5px solid black;">' . strtoupper(lang('sale_price_orig_short')) . '</td>
                            <td bgcolor="#eaeaea" style="font-size:6pt; width: 18%;text-align: right; border-right: 0.5px solid black; border-bottom: 0.5px solid black;">' . number_format($sale_price_orig, 2, ',', '.') . '</td>
                        </tr>';
        }

        $html .=    '</tbody>
                </table>';

        return $html;
    }

    public function view_sub_details($subdetails, $father_code, &$html = ''){

        $counter        = 1;
        $tot_subdetails = count($subdetails);
        foreach ($subdetails as $subdetail) {
            if (array_key_exists('subdetails', $subdetail)) {
                $tot_subdetails += count($subdetail['subdetails']);
            }
        }

        foreach ($subdetails as $subdetail) {
            $row_span = 1 + ((array_key_exists('subdetails', $subdetail)) ? count($subdetail['subdetails']) : 0);
            $line = '';
            if($counter == $tot_subdetails){
                //$line = 'border-bottom: solid 2px darkgrey!important;';
            }

            $html .= " <tr class=\"row_details comp_row " . $father_code . "\" style=\"" . $line . "\">";
            if($subdetail['is_purchase']){
                $html .= "
                        <td class=\"f_calibri thick_border_left thick_border_right\">" . $subdetail['code'] . "</td>
                        <td class=\"f_calibri thick_border_left thick_border_right\">" . $subdetail['description'] . "</td>
                        <td class=\"f_calibri text-right qty_col thick_border_left \"></td>
                        <td class=\"f_calibri text-right weight_col \">" . number_format($subdetail['qty'], 4, ',', '.') . "</td>
                        <td class=\"f_calibri text-right unit_cost_col\">" . number_format($subdetail['cost'], 4, ',', '.') . "</td>
                        <td class=\"f_calibri text-right unit_cost_col thick_border_right\">" . number_format($subdetail['cost'] * $subdetail['qty']/* * $subdetail['weight']*/, 2, ',', '.') . "</td>
                        <td class=\"f_calibri thick_border_left thick_border_right\" colspan=\"4\"></td>
                    </tr>";
            }
            else{
                $html .= "
                        <td class=\"f_calibri italic f_lightgrey_bg               bold grey_border_top             thick_border_left thick_border_right\">" . $subdetail['code'] . "</td>
                        <td class=\"f_calibri italic f_lightgrey_bg               bold grey_border_top             thick_border_left thick_border_right\">" . $subdetail['description'] . "</td>
                        <td class=\"f_calibri italic f_lightgrey_bg qty_col       grey_border_top text-right  thick_border_left\">" . number_format($subdetail['qty'], 0) . "</td>
                        <td class=\"f_calibri italic f_lightgrey_bg               grey_border_top text-right  thick_border_right\" colspan=\"3\"></td>
                        <td class=\"f_calibri italic f_lightgrey_bg weight_col    grey_border_top text-right  thick_border_left\" >" . $this->CI->utility->decimal_to_hour_min_sec($subdetail['theoric_time']) . "</td>
                        <td class=\"f_calibri italic f_lightgrey_bg weight_col    grey_border_top text-right\"                   >" . number_format($subdetail['effect_percentage'], 2, ',', '.'). " %</td>
                        <td class=\"f_calibri italic f_lightgrey_bg unit_cost_col grey_border_top text-right\"                   >" . number_format($subdetail['hour_cost'], 2, ',', '.') . "</td>
                        <td class=\"f_calibri italic f_lightgrey_bg prod_cost_col grey_border_top text-right  thick_border_right\">" . number_format($subdetail['production_cost'], 2, ',', '.') ."</td>
                        <td class=\"f_calibri                                         grey_border_top text-right  thick_border_left   text-mid\" rowspan=\"" . $row_span . "\">" . number_format($subdetail['cost'] * $subdetail['qty'], 2, ',', '.') ."</td>                   
                    </tr>";
            }

            if (!empty($subdetail['subdetails'])) {
                $this->view_sub_details($subdetail['subdetails'], $subdetail['code'], $html);
            }

            $counter++;
        }

        return $html;
    }



    // INIT OFF6985
    public function export_excell_cost_card($product_cost_card_data) {

        //load our new PHPExcel library
        $this->CI->load->library('excel');

        $xls_filename = 'Costo prodotto [' . $product_cost_card_data['item_code'] . '] ' . $product_cost_card_data['item_description'];

        $this->CI->excel->setActiveSheetIndex(0);
        //name the worksheet
        $this->CI->excel->getActiveSheet()->setTitle(strtoupper(lang('menu_roles')));

        /*************************************************************/
        $this->compose_excel($product_cost_card_data);
        /*************************************************************/

        $this->CI->excel->setActiveSheetIndex(0);
        $filename=  $xls_filename . '_' . date('dmY'); //save our workbook as this file name
        @ob_end_clean();
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=$filename.xlsx");
        header("Content-Transfer-Encoding: binary");
        $objWriter = PHPExcel_IOFactory::createWriter($this->CI->excel, 'Excel2007');
        $objWriter->save('php://output');
    }
    var $style = array(
        'custom_blue' => '0F4E74',
        'header_td'   => '1B8FD3',
        'bg_separator'   => 'EEEEEE',
        'f_lightgrey_bg'   => 'eaeaea'
    );

    var $xlsRow = 0;

    var $xlsTree = array();
    var $xlsRowValue = array();
    var $customBorder = array();
    public function compose_excel($product_cost_card_data) {

        $portlet_title_style_top = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => 'cccccc')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );

        $portlet_title_style_bottom = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['custom_blue'])
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => 'FFFFFF')
                )
            ),
            'font' => array(
                'color' => array('rgb' => 'FFFFFF')
            )
        );
        $separator = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['bg_separator'])
            ),
        );
        $standard_td = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '1B8FD3')
            ),
            'borders' => array(
                'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000'))
            ),
            'font' => array(
                'color' => array('rgb' => 'FFFFFF'),
                'bold' => true,
            )
        );
        $this->fill_white = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            )
        );
        $this->standard_td_border_bottom_none = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '1B8FD3')
            ),
            'borders' => array(
                'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000'))
            ),
            'font' => array(
                'color' => array('rgb' => 'FFFFFF'),
                'bold' => true,
            )
        );
        $this->standard_td_border_top_white = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '1B8FD3')
            ),
            'borders' => array(
                'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => 'FFFFFF')),
                'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => 'FFFFFF')),
                'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => 'FFFFFF'))
            ),
            'font' => array(
                'color' => array('rgb' => 'FFFFFF'),
                'bold' => true,
            )
        );
        $this->total_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'eaeaea')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            ),
            'font' => array(
                'bold' => true,
            )
        );
        $this->customBorder = array(
            'ea_top_black_right_black_left_black' => array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000'))
                )
            ),
            'ea_top_no_right_black_left_black' => array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000'))
                )
            ),
            'ea_top_black_right_black'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'ea_top_none_right_black'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'ea_top_black_right_black_font_bold'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('rgb' => '000000')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                ),
                'font' => array(
                    'bold' => true,
                )
            ),
            'ea_top_black_right_cccccc' => array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc'))
                )
            ),
            'ea_top_none_right_cccccc' => array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc'))
                )
            ),
            'nobg_top_cccccc_right_black_left_black' => array(
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'nobg_top_cccccc_right_black' => array(
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'nobg_top_cccccc_right_black_font_bold' => array(
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                ),
                'font' => array(
                    'bold' => true,
                )
            ),
            'nobg_top_cccccc_right_cccccc'=> array(
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc'))
                )
            ),
            'nobg_top_cccccc'=> array(
                'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')))
            ),
            'eadg_top_dg_right_black_left_black'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN , 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'eadg_top_dg_right_black_left_black_bold'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                ),
                'font' => array(
                    'bold' => true,
                )
            ),
            'eadg_top_dg_right_black'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'eadg_top_dg_right_black_font_bold'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                ),
                'font' => array(
                    'bold' => true,
                )
            ),
            'eadg_top_dg_right_black_font_bold_italic'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                ),
                'font' => array(
                    'bold' => true,
                    'italic' => true,
                )
            ),
            'eadg_top_dg_right_cccccc'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc'))
                )
            ),
            'eadg_top_dg'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')))
            ),
            'eadg_top_dg_right_black_left_black_fi'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')),
                    'left' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'eadg_top_dg_right_black_fi'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000'))
                )
            ),
            'eadg_top_dg_right_cccccc_fi'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array(
                    'top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')),
                    'right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'cccccc')))
            ),
            'eadg_top_dg_fii'=> array(
                'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'eaeaea')),
                'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => 'a9a9a9')))
            ),
            'nobg_right_black'=> array(
                'borders' => array('right' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')))
            ),
            'nobg_top_black'=> array(
                'borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')))
            ),
            'nobg_left_black'=> array(
                'borders' => array('left' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')))
            ),
            'nobg_bottom_black'=> array(
                'borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('rgb' => '000000')))
            ),
        );

        $column = 0;
        $row = 1;
        $this->xlsRow = 1;
        $this->xlsColumns = array(
            'direct_costs' => 0,
            'direct_costs_string' => PHPExcel_Cell::stringFromColumnIndex(0),
            'productions' => 1,
            'productions_string' => PHPExcel_Cell::stringFromColumnIndex(1),
            'materials' => array('start' => 2, 'end' => 5),
            'materials_string' => array('start' => PHPExcel_Cell::stringFromColumnIndex(2), 'end' => PHPExcel_Cell::stringFromColumnIndex(5)),
            'times' => array('start' =>6, 'end' => 9),
            'times_string' => array('start' => PHPExcel_Cell::stringFromColumnIndex(6), 'end' => PHPExcel_Cell::stringFromColumnIndex(9)),
            'totals' => 10,
            'totals_string' => PHPExcel_Cell::stringFromColumnIndex(10),
            'pz_short' => 2,
            'pz_short_string' => PHPExcel_Cell::stringFromColumnIndex(2),
            'weight' => 3,
            'weight_string' => PHPExcel_Cell::stringFromColumnIndex(3),
            'kg_cost' => 4,
            'kg_cost_string' => PHPExcel_Cell::stringFromColumnIndex(4),
            'total' => 5,
            'total_string' => PHPExcel_Cell::stringFromColumnIndex(5),
            'label_time' => 6,
            'label_time_string' => PHPExcel_Cell::stringFromColumnIndex(6),
            '_ineff_short' => 7,
            '_ineff_short_string' => PHPExcel_Cell::stringFromColumnIndex(7),
            'h_cost_shorter' => 8,
            'h_cost_shorter_string' => PHPExcel_Cell::stringFromColumnIndex(8),
            'time_total' => 9,
            'time_total_string' => PHPExcel_Cell::stringFromColumnIndex(9)
        );
        //$currentColumn = 'A';
        //$columnIndex = PHPExcel_Cell::columnIndexFromString($currentColumn);
       //$adjustedColumn = PHPExcel_Cell::stringFromColumnIndex($adjustedColumnIndex - 1);

        //PORTLET TITLE
        $portlet_title = '[' . $product_cost_card_data['item_code'] . '] - ' .  $product_cost_card_data['item_description'] . '] - ' . date('d/m/Y', strtotime(is_null($product_cost_card_data['date_modify']) ? $product_cost_card_data['date_insert'] : $product_cost_card_data['date_modify']));
        $this->CI->excel->getActiveSheet()->getStyle('A1:L1')->applyFromArray($portlet_title_style_top);
        $this->CI->excel->getActiveSheet()->mergeCells("A" . $this->xlsRow . ":L" . $this->xlsRow);
        $this->CI->excel->getActiveSheet()->getRowDimension(1)->setRowHeight(30);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow("A1")->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $portlet_title);
        $row++;
        $this->xlsRow++;
        //$this->CI->excel->getActiveSheet()->getStyle("A2:L2")->applyFromArray($portlet_title_style_bottom);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0, 10)->setWidth(25);
        //$this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(30);
        $row++;
        $this->xlsRow++;
        $this->CI->excel->getActiveSheet()->getStyle("A3:F3")->applyFromArray($separator);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0, 10)->setWidth(50);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(3);
        $row++;
        $this->xlsRow++;
        /*COMPONENTI DI PRODUZIONE*/
        $prod_components_res = $this->export_prod_components($product_cost_card_data,$standard_td, $row);
        $tot_prod_components = $prod_components_res['tot_prod_components'];
        /*ACUISTI*/
        $purc_components_res = $this->export_acq_components($product_cost_card_data,$standard_td);
        $tot_purc_components = $purc_components_res['tot_purc_components'];
        /*COSTI ASSEMBLAGGIO*/
        $asse_components_res = $this->export_assembly_cost($product_cost_card_data,$standard_td);
        $tot_asse_components = $asse_components_res['tot_asse_components'];
        /*COSTI INDIRETTI*/
        $indi_components_res =  $this->export_indirect_cost($product_cost_card_data,$standard_td);
        $tot_indi_components = $indi_components_res['tot_indi_components'];


        /*TOTALE*/
        $tot_no_taxes       = $tot_prod_components + $tot_purc_components + $tot_asse_components;
        $sale_price         = $this->CI->utility->get_field_value('sale_price', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $sale_price_orig    = $this->CI->utility->get_field_value('sale_price_orig', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $taxes              = $this->CI->utility->get_field_value('taxes', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
        $year               = $this->CI->utility->get_field_value('year', $this->CI->db->tb_article_cost, array('id_article_cost' => $product_cost_card_data['id_article_cost']));
        $cost_config_perc   = $this->CI->utility->get_field_value('tax_perc', $this->CI->db->tb_tab_cost_config, array('year' => $year));
        $taxes              = ($sale_price - ($tot_no_taxes + $tot_indi_components)) * $cost_config_perc/100;
        if($taxes < 0){
            $taxes = 0;
        }
        $this->export_total_cost($product_cost_card_data, $tot_no_taxes, $tot_indi_components, $taxes, $sale_price, $sale_price_orig, $standard_td,$cost_config_perc);


    }
    var $test_start_row = 0;
    public function export_prod_components($product_cost_card_data,$standard_td, $row) {
        $prod_components_submai_style3 = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $total_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'eaeaea')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            ),
            'font' => array(
                'bold' => true,
            )
        );
        $this->export_prod_components_header($standard_td, $row);
        $row++;
        $row++;
        $this->xlsRow++;
        $this->xlsRow++;
        $this->test_start_row = $this->xlsRow;
        $tot_prod_components = 0;

        if(!empty($product_cost_card_data['production_components_cost'])) {
            if(!empty($product_cost_card_data['production_components_cost']['production_components'])) {
                //foreach ($product_cost_card_data['production_components_cost']['production_components'] as $production_component) {
                $counter = count($product_cost_card_data['production_components_cost']['production_components']);
                for ($i = 0; $i < $counter; $i++) {
                    if(substr($product_cost_card_data['production_components_cost']['production_components'][$i]['code'], 0, 3) != '60.') {
                        $tot_prod_components += $product_cost_card_data['production_components_cost']['production_components'][$i]['component_cost'] * $product_cost_card_data['production_components_cost']['production_components'][$i]['qty'];
                        $this->export_prod_components_main($product_cost_card_data['production_components_cost']['production_components'][$i],$this->xlsRow);
                        $row++;
                        if(!empty($product_cost_card_data['production_components_cost']['production_components'][$i]['details'])) {
                            $detail_counter = count($product_cost_card_data['production_components_cost']['production_components'][$i]['details']);
                            for($j = 0; $j < $detail_counter; $j++) {
                                $detail = $product_cost_card_data['production_components_cost']['production_components'][$i]['details'][$j];
                                $this->export_prod_components_submain($detail,$this->xlsRow, $product_cost_card_data['production_components_cost']['production_components'][$i]['id']);
                                $row++;
                                $this->export_prod_components_subs($detail,$row, $product_cost_card_data['production_components_cost']['production_components'][$i]['id']);
                                $row++;
                            }
                        }
                    }
                    if(!empty($this->xlsRowValue['sub_tot_row_span'])) {
                        $end = end($this->xlsRowValue['sub_tot_row_span']);
                        $mid = intval(str_replace("K", '',$end))-intval(str_replace("K", '',$this->xlsRowValue['sub_tot_row_span'][0]));
                        $totRow = round($mid/2)+intval(str_replace("K", '',$this->xlsRowValue['sub_tot_row_span'][0]));
                        $this->CI->excel->getActiveSheet()->getStyle($this->xlsRowValue['sub_tot_row_span'][0] . ":" .$end)->applyFromArray($prod_components_submai_style3);
                        $this->CI->excel->getActiveSheet()->getStyle(str_replace("K","L",$this->xlsRowValue['sub_tot_row_span'][0]) . ":" .str_replace("K","L",$end))->applyFromArray($prod_components_submai_style3);
                        $this->CI->excel->getActiveSheet()->getStyle("L".$totRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                        $this->CI->excel->getActiveSheet()->setCellValue("L".$totRow, $this->xlsRowValue['subtotale']);
                    }
                }
            }
        }
        $this->CI->excel->getActiveSheet()->getStyle('L'. $this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        $end = (!empty($this->xlsRowValue['produzioni']['L']) && end($this->xlsRowValue['produzioni']['L'])) ? end($this->xlsRowValue['produzioni']['L']) : '6';
        $this->CI->excel->getActiveSheet()->setCellValue('L'. $this->xlsRow,'=SUM(L6:L' . $end .')');
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow('K',$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow('L',$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L"  . $this->xlsRow)->applyFromArray($total_style);
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":J"  . $this->xlsRow)->applyFromArray($prod_components_submai_style3);
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow+1) . ":L"  . ($this->xlsRow+1))->applyFromArray($this->fill_white);
        $this->xlsRowValue['riepilogo_produzione'] = $this->xlsRow;
        $result['tot_prod_components']  = $tot_prod_components;

            //$this->CI->excel->getActiveSheet()->getStyle('K' . $this->xlsRowValue['produzioni']['L'][0] . ':L' .($this->xlsRowValue['produzioni_componente']['L'][0]-1))->applyFromArray($total_style);
            //$this->CI->excel->getActiveSheet()->mergeCells("K" . $this->xlsRowValue['produzioni']['L'][0] . ":L" . ($this->xlsRowValue['produzioni_componente']['L'][0]-1));
        $this->xlsRow++;
        return $result;

    }
    var $caluculate = array();
    var $formula = array();
    public function export_prod_calculate_value($main) {
        $n = 0;
        if(!empty($main)) {
            foreach($main as $key => $value) {
                if (substr($value['code'], 0, 3) != '60.') {
                    $this->caluculate[$value['id']]['main'] = $this->test_start_row;
                    $this->test_start_row++;
                    if (!empty($value['details'])) {
                        $detail_counter = count($value['details']);
                        for ($i = 0; $i < $detail_counter; $i++) {
                            $this->export_prod_calculate_detail_value($value['details'][$i], $value['id']);
                            $this->export_prod_calculate_subdetail_value($value['details'][$i], $value['id']);
                        }
                    }
                }
                $n++;
            }
        }
    }

    public function export_prod_calculate_detail_value($details, $id) {
        //SCRIVO RIGA DETAIL
        if($details['is_purchase']){
            $this->caluculate[$id]['detail_purchase'][$details['code']] = $this->test_start_row;
            $this->test_start_row++;
        } else {
            $this->caluculate[$id]['detail_nopurchase'][$details['code']] = $this->test_start_row;
            $this->test_start_row++;
        }
    }

    public function export_prod_calculate_subdetail_value($details,$id) {
        if (array_key_exists('subdetails', $details) && !empty($details['subdetails'])) {
            $subdetail_counter = count($details['subdetails']);
            for($k = 0; $k < $subdetail_counter; $k++) {
                $subdetail  = $details['subdetails'][$k];
                if($subdetail['is_purchase']){
                    $this->caluculate[$id]['subdetail_purchase'][$subdetail['code']] = $this->test_start_row;
                    $this->test_start_row++;
                } else {
                    $this->caluculate[$id]['subdetail_nopurchase'][$subdetail['code']] = $this->test_start_row;
                    $this->test_start_row++;
                }
                //SCRIVO RIGA DETAIL

                $this->export_prod_calculate_subdetail_value($subdetail,$id);
            }
        }
    }

    public function export_prod_components_header($tandard_td, $row) {

        //COSTI DIRETTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("A" . $row . ":A" . (intval($this->xlsRow)+1))->applyFromArray($tandard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("A" . $row . ":A" . (intval($this->xlsRow+1)));
        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(35);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow("A1")->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP );
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, strtoupper(lang('direct_costs')));
        //PRODUZIONI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("B" . $row . ":B" . (intval($this->xlsRow+1)))->applyFromArray($tandard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("B" . $row . ":B" . (intval($this->xlsRow+1)));
        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(40);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$row)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('productions')));
        //MATERIALI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow . ":F" . $this->xlsRow)->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("C" . $this->xlsRow . ":F" . $this->xlsRow);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow("C".$row)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['materials']['start'],$this->xlsColumns['materials']['end'])->setWidth(80);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['materials']['start'], $this->xlsRow, strtoupper(lang('materials')));
        //MATERIALI RIGHE
        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(8);
        $this->CI->excel->getActiveSheet()->getStyle("C" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['materials']['start'], (intval($this->xlsRow+1)), strtoupper(lang('pz_short')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['weight'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("D" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['materials']['start']+1), (intval($this->xlsRow+1)), strtoupper(lang('weight')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['kg_cost'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("E" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['materials']['start']+2), (intval($this->xlsRow+1)), strtoupper(lang('kg_cost')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['total'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("F" . (intval($row+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['materials']['start']+3), (intval($this->xlsRow+1)), strtoupper(lang('total')));
        //TEMPI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("G"  . $this->xlsRow . ":J" .  $this->xlsRow);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow("C".$row)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['times']['start'],$this->xlsColumns['times']['end'])->setWidth(80);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['times']['start'],$this->xlsRow, strtoupper(lang('times')));

        //TEMPI RIGHE
        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['label_time'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("G" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['times']['start'], (intval($this->xlsRow+1)), strtoupper(lang('label_time')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['label_time'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("H" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['times']['start']+1), (intval($this->xlsRow+1)), strtoupper(lang('%_ineff_short')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['h_cost_shorter'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("I" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['times']['start']+2), (intval($this->xlsRow+1)), strtoupper(lang('h_cost_shorter')));

        $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['time_total'])->setWidth(12);
        $this->CI->excel->getActiveSheet()->getStyle("J" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['times']['start']+3), (intval($this->xlsRow+1)), strtoupper(lang('total')));

        //TOTALE HEAD
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($tandard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,(intval($this->xlsRow)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,(intval($this->xlsRow)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['totals'], $this->xlsRow, strtoupper(lang('total')));
    }

    public function export_prod_components_main($elem,$row) {
        $tot_row_span       = 1 + count($elem['details']);
        $subtot_row_span    = 1;
        $subtot_acq_comp    = 0;
        $subtot_acq_comp_formula = '';
        foreach ($elem['details'] as $detail) {
            if(!empty($detail['subdetails'])){
                $tot_row_span += count($detail['subdetails']);
                foreach ($detail['subdetails'] as $subdetail) {
                    if(!$subdetail['is_purchase']){
                        $tot_row_span++;
                    }
                }
            }
            if($detail['is_purchase']){
                $subtot_row_span++;
                $subtot_acq_comp += $detail['cost'] *$detail['qty'] * $detail['weight'];

            }
        }

        //COSTI DIRETTI CODICE
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_no_right_black_left_black']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $elem['code']);
        //PRODUZIONI DESCRIZIONE
        //$this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_black_right_black']);
        $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_no_right_black_left_black']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($elem['description'], 0, 30));
        //MATERIALI PZ
        $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $qty = intval($elem['qty']);
        $this->xlsRowValue[$elem['id']][] = $this->xlsRow;
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow,$qty);
        //$this->CI->excel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(2).$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
        //MATERIALI PESO
        $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, '');
        //MATERIALI COSTO KG
        $this->CI->excel->getActiveSheet()->getStyle("E" . $this->xlsRow )->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(4,$this->xlsRow, '');
        //MATERIALI TOTALE
        $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_black']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(5,$this->xlsRow, '');


        //TEMPI TEMPO
        $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow )->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, number_format($elem['theoric_time'],2,',', ''));
        //TEMPI INEFF.
        $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow, number_format($elem['effect_percentage'],2,',', ''));
        //TEMPI COSTO ORA
        $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_cccccc']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, number_format($elem['hour_cost'],2,',', ''));
        //TEMPI TOTALE

        $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow)->applyFromArray($this->customBorder['ea_top_none_right_black']);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyle("J".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        $this->CI->excel->getActiveSheet()->setCellValue("J".$this->xlsRow, "=G" . $this->xlsRow ."*60*(1+H" . $this->xlsRow ."/100)*I" . $this->xlsRow . "/3600");

        if($tot_row_span > $subtot_row_span){
            $this->xlsRowValue['tot_row_span'] = 1;
            //$this->CI->excel->getActiveSheet()->setCellValue('K'.$this->xlsRow,'=C6*' .($elem['production_cost'] + $subtot_acq_comp));
            /*$this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue('L'.$this->xlsRow, $elem['component_cost'] * $elem['qty']);*/
            $this->xlsRowValue['subtotale'] = $elem['component_cost'] * $elem['qty'];
            $this->xlsRowValue['produzioni_costo']['k'][] = $this->xlsRow;
            $this->xlsRowValue['produzioni_componente']['L'][] = $this->xlsRow;
        }
        else{
            $this->xlsRowValue['tot_row_span'] = 2;
            //$this->CI->excel->getActiveSheet()->setCellValue('L'. $this->xlsRow,'=C' . $this->xlsRow . '*' .$elem['component_cost']);
            $this->xlsRowValue['produzioni']['L'][] = $this->xlsRow;
        }
        //$this->xlsRowValue['main'][$elem['code']][$elem['id']] = $this->xlsRow;
        $this->xlsRow++;

    }

    public function export_prod_components_submain($detail, $row, $id) {
        $prod_components_submai_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $prod_components_submai_style2 = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $subdetail_row_span     = 1;
        $subtot_acq_subdetails  = 0;
        if(!empty($detail['subdetails'])){
            foreach ($detail['subdetails'] as $subdetail) {
                if($subdetail['is_purchase']){
                    $subdetail_row_span++;
                    $subtot_acq_subdetails += $subdetail['cost'] * $subdetail['qty'] /** $subdetail['weight']*/;
                }
            }
        }

        if($detail['is_purchase']){
            //COSTI DIRETTI CODICE
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black_left_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $detail['code']. '');
            //PRODUZIONI DESCRIZIONE
            $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($detail['description'], 0, 30));
            //MATERIALI PZ
            $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');
            //MATERIALI PESO
            $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, number_format($detail['qty'], 4));
            //MATERIALI COSTO KG
            $this->CI->excel->getActiveSheet()->getStyle("E" . $this->xlsRow )->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(4,$this->xlsRow, number_format($detail['cost'], 4));
            //MATERIALI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("F".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("F".$this->xlsRow, "=D" . $this->xlsRow . "*E" . $this->xlsRow);

            //TEMPI TEMPO
            $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow )->applyFromArray($this->customBorder['nobg_top_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, '');
            //TEMPI INEFF.
            $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc']);
           // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow, '');
            //TEMPI COSTO ORA
            $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, '');
            //TEMPI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(9,$this->xlsRow, '');

            //TOTALE TOTALE+TOTALE DETAIL
            if ($this->xlsRowValue['tot_row_span'] == 1) {
                $this->CI->excel->getActiveSheet()->getStyle("K" . ($this->xlsRow-1))->applyFromArray($prod_components_submai_style);
                $this->CI->excel->getActiveSheet()->getStyle("L" . ($this->xlsRow-1))->applyFromArray($prod_components_submai_style);
                $this->CI->excel->getActiveSheet()->getStyle("K" . ($this->xlsRow))->applyFromArray($prod_components_submai_style2);
                $this->CI->excel->getActiveSheet()->getStyle("L" . ($this->xlsRow))->applyFromArray($prod_components_submai_style2);
                $this->CI->excel->getActiveSheet()->getStyle("K".($this->xlsRow-1))->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                $this->CI->excel->getActiveSheet()->setCellValue("K".($this->xlsRow-1), "=(C" . ($this->xlsRow-1) . "*F" . $this->xlsRow . ")+(C" . ($this->xlsRow-1) . "*J" . ($this->xlsRow-1) . ")");
                $this->xlsRowValue['sub_tot_row_span'][] = "K".($this->xlsRow);
            } elseif($this->xlsRowValue['tot_row_span'] == 2) {
                $this->CI->excel->getActiveSheet()->getStyle("K" . ($this->xlsRow-1) . ":L" .$this->xlsRow)->applyFromArray($prod_components_submai_style);
                $this->CI->excel->getActiveSheet()->getStyle("L".($this->xlsRow-1))->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                $this->CI->excel->getActiveSheet()->setCellValue("L".($this->xlsRow-1), "=(C" . ($this->xlsRow-1) . "*F" . $this->xlsRow . ")+(C" . ($this->xlsRow-1) . "*J" . ($this->xlsRow-1) . ")");
            }
            //$this->xlsRowValue['deatail_purchaise']['K'][] = $this->xlsRow;
            //$this->xlsRowValue['deatail_purchaise'][$detail['code']][$id] = $this->xlsRow;//COSTRUISCO ALBERO PER CALCOLI

        } else {
            //COSTI DIRETTI CODICE
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black_left_black_bold']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $detail['code'] . '');
            //PRODUZIONI DESCRIZIONE
            //$this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
            $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black_font_bold']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($detail['description'], 0, 30));
            //MATERIALI PZ
            $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, number_format($detail['qty'], 4));
            //MATERIALI PESO
            $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, '' );
            //MATERIALI COSTO KG
            $this->CI->excel->getActiveSheet()->getStyle("E" . $this->xlsRow )->applyFromArray($this->customBorder['eadg_top_dg']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth(10);
            //$this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(4,$this->xlsRow, '' );
            //MATERIALI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(5,$this->xlsRow, '');


            //MATERIALI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow,  number_format($detail['theoric_time'], 2, ',',''));
            //TEMPI TEMPO
            $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow )->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow,  number_format($detail['effect_percentage'], 2, ',',''));
            //TEMPI INEFF.
            $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, number_format($detail['hour_cost'], 2, ',',''));
            //TEMPI COSTO ORA
            $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("J".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("J".$this->xlsRow, "=G" . $this->xlsRow ."*60*(1+H" . $this->xlsRow ."/100)*I" . $this->xlsRow . "/3600");
            //TEMPI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_right_black']);
            $this->CI->excel->getActiveSheet()->getStyle("K". $this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("K". $this->xlsRow,'=C' . $this->xlsRow. '*' .($detail['production_cost'] + $subtot_acq_subdetails));
            $this->xlsRowValue['sub_tot_row_span'][] = "K".($this->xlsRow);
            //$this->xlsRowValue['deatail_no_purchaise']['K'][] = $this->xlsRow;
            //$this->xlsRowValue['deatail_no_purchaise'][$detail['code']][$id] = $this->xlsRow;//COSTRUISCO ALBERO PER CALCOLI

        }
        $this->xlsRow++;
    }

    public function export_prod_components_subs($detail, $row, $id = null) {

        if (array_key_exists('subdetails', $detail) && !empty($detail['subdetails'])) {
            $subdetail_counter = count($detail['subdetails']);
            for($k = 0; $k < $subdetail_counter; $k++){
                $subdetail  = $detail['subdetails'][$k];
                if($subdetail['is_purchase']){
                    //COSTI DIRETTI CODICE
                    $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black_left_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $subdetail['code'] . '');
                    //PRODUZIONI DESCRIZIONE
                    $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($subdetail['description'], 0, 30));
                    //MATERIALI PZ
                    $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');
                    //MATERIALI PESO
                    $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, number_format($subdetail['qty'], 2, ',', ''));
                    //MATERIALI COSTO KG
                    $this->CI->excel->getActiveSheet()->getStyle("E" . $this->xlsRow )->applyFromArray($this->customBorder['nobg_top_cccccc_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(4,$this->xlsRow, number_format($subdetail['cost'], 2, ',', ''));
                    //MATERIALI TOTALE
                    $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getStyle("F".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $this->CI->excel->getActiveSheet()->setCellValue("F".$this->xlsRow, "=D" . $this->xlsRow . "*E" . $this->xlsRow);


                    //TEMPI TEMPO
                    $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow )->applyFromArray($this->customBorder['nobg_top_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, '');
                    //TEMPI INEFF.
                    $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow, '');
                    //TEMPI COSTO ORA
                    $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, '');
                    //TEMPI TOTALE
                    $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_top_cccccc_right_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(9,$this->xlsRow, '');
                    //$this->xlsRowValue['deatail_sub_purchaise']['K'][] = $this->xlsRow;
                    //$this->xlsRowValue['subdeatail_purchaise'][$subdetail['code']][$id] = $this->xlsRow;//COSTRUISCO ALBERO PER CALCOLI

                    $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_right_black']);
                    $this->xlsRowValue['sub_tot_row_span'][] = "K".($this->xlsRow);
                } else {
                    //COSTI DIRETTI CODICE
                    //$this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black_left_black']);
                    $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black_left_black_bold']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $subdetail['code']  . '');
                    //PRODUZIONI DESCRIZIONE
                    //$this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
                    $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black_font_bold_italic']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($subdetail['description'], 0, 30));
                    //MATERIALI PZ
                    $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, number_format($subdetail['qty'], 0));
                    //MATERIALI PESO
                    $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, '');
                    //MATERIALI COSTO KG
                    $this->CI->excel->getActiveSheet()->getStyle("E" . $this->xlsRow )->applyFromArray($this->customBorder['eadg_top_dg']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(4)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(4,$this->xlsRow, '');
                    //MATERIALI TOTALE
                    $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(5,$this->xlsRow, '');
                    //TEMPI TEMPO
                    $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow )->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, number_format($subdetail['theoric_time'], 2, ',', ''));
                    //TEMPI INEFF.
                    $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow, number_format($subdetail['effect_percentage'], 2, ',', ''));
                    //TEMPI COSTO ORA
                    $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_cccccc']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, number_format($subdetail['hour_cost'], 2, ',', ''));
                    //TEMPI TOTALE
                    $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow)->applyFromArray($this->customBorder['eadg_top_dg_right_black']);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

                    $this->CI->excel->getActiveSheet()->getStyle("J".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $this->CI->excel->getActiveSheet()->setCellValue("J".$this->xlsRow, "=G" . $this->xlsRow ."*60*(1+H" . $this->xlsRow ."/100)*I" . $this->xlsRow . "/3600");

                    $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow)->applyFromArray($this->customBorder['nobg_right_black']);

                    $this->CI->excel->getActiveSheet()->getStyle("K". $this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $this->CI->excel->getActiveSheet()->setCellValue("K". $this->xlsRow,'=C'. $this->xlsRow .'*' .$subdetail['cost']);
                    $this->xlsRowValue['sub_tot_row_span'][] = "K".($this->xlsRow);
                    //$this->xlsRowValue['deatail_no_sub_purchaise']['K'][] = $this->xlsRow;
                    //$this->xlsRowValue['subdeatail_no_purchaise'][$subdetail['code']][$id] = $this->xlsRow;//COSTRUISCO ALBERO PER CALCOLI
                }
                $row++;
                $this->xlsRow++;
                $this->export_prod_components_subs($subdetail,$row);
            }
        }
    }

    public function export_acq_components($product_cost_card_data,$standard_td) {

        $this->export_acq_components_header($standard_td);

        $res = $this->export_acq_components_body($product_cost_card_data,$standard_td);

        $result['tot_purc_components']  = $res['tot_purc_components'];

        return $result;
    }

    public function export_acq_components_header($tandard_td) {
        $this->xlsRow++;

        //COSTI DIRETTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . (intval($this->xlsRow)+1))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("A" . $this->xlsRow . ":A" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow("A1")->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow )->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, strtoupper(lang('direct_costs')));
        //ACQUISTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)));
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow )->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('purchases')));
        //MATERIALI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow . ":F" . $this->xlsRow)->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("C" . $this->xlsRow . ":F" . $this->xlsRow);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow )->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['materials']['start'],$this->xlsColumns['materials']['end'])->setWidth(80);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['materials']['start'], $this->xlsRow, strtoupper(lang('materials')));
        //MATERIALI QTA HEAD
        $this->CI->excel->getActiveSheet()->getStyle("C" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,(intval($this->xlsRow+1)))->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['materials']['start'], (intval($this->xlsRow+1)), strtoupper(lang('common_qty')));
        //MATERIALI COSTO UNITARIO
        $this->CI->excel->getActiveSheet()->getStyle("D" . (intval($this->xlsRow+1)). ":E" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        $this->CI->excel->getActiveSheet()->mergeCells("D" . (intval($this->xlsRow+1)) . ":E" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['weight'])->setWidth(40);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,(intval($this->xlsRow+1)))->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['materials']['start']+1), (intval($this->xlsRow+1)), strtoupper(lang('unit_cost_short')));
        //MATERIALI TOTALE
        $this->CI->excel->getActiveSheet()->getStyle("F" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['total'])->setWidth(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,(intval($this->xlsRow+1)))->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(4,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(intval($this->xlsColumns['materials']['start']+3), (intval($this->xlsRow+1)), strtoupper(lang('total')));

        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  intval(($this->xlsRow+1)))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("G"  . $this->xlsRow . ":J" .  intval(($this->xlsRow+1)));
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,(intval($this->xlsRow+1)))->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,(intval($this->xlsRow+1)))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,(intval($this->xlsRow+1)))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['times']['start'],$this->xlsColumns['times']['end'])->setWidth(80);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['times']['start'],$this->xlsRow, '');


        //TOTALE HEAD
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($tandard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['totals'], $this->xlsRow, strtoupper(lang('total')));
        $this->xlsRow++;
        $this->xlsRow++;

    }

    public function export_acq_components_body($product_cost_card_data,$standard_td ) {
        $html = '';
        $mid_categories_acq = array();
        $puchase_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            /*'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $puchase_style_white_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            /*'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $costi_indiretti_code_style = array(
            'borders' => array(
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
        );
        $puchase_white_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
        );
        if(!empty($product_cost_card_data['purchase_components_cost'])) {
            if (!empty($product_cost_card_data['purchase_components_cost']['purchase_components'])) {
                foreach ($product_cost_card_data['purchase_components_cost']['purchase_components'] as $purchase_components) {
                    $current_mid_category = explode('.', $purchase_components['code'])[1];
                    if(!array_key_exists($current_mid_category, $mid_categories_acq)){
                        $mid_categories_acq[$current_mid_category]['tot_rows']   = 0;
                        $mid_categories_acq[$current_mid_category]['sub_tot']    = 0;
                    }
                    $mid_categories_acq[$current_mid_category]['tot_rows']++;
                    $mid_categories_acq[$current_mid_category]['sub_tot'] += $purchase_components['qty'] * $purchase_components['cost'];
                }
            }
        }

        $tot_purc_components = 0;
        $firstRow = $this->xlsRow;
        if(!empty($product_cost_card_data['purchase_components_cost'])) {
            if(!empty($product_cost_card_data['purchase_components_cost']['purchase_components'])) {
                $counter = 0;
                $old_mid_category = false;
                foreach ($product_cost_card_data['purchase_components_cost']['purchase_components'] as $purchase_components) {
                    $bg_color   = $puchase_style;
                    $border_array = array(
                        'borders' => array(
                            'top' => array(
                                'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                'color' => array('rgb' => 'cccccc')
                            )
                        )
                    );
                    $border_total_array = array(
                        'borders' => array(
                            'top' => array(
                                'style' => PHPExcel_Style_Border::BORDER_THIN,
                                'color' => array('rgb' => 'FFFFFF')
                            )
                        )
                    );
                    if($counter%2 != 0){
                        $bg_color   = $puchase_style_white_style;
                    }
                    $current_mid_category = explode('.', $purchase_components['code'])[1];
                    if($old_mid_category != $current_mid_category){
                        $border_array = array(/*
                            'borders' => array(
                                'top' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                    'color' => array('rgb' => 'c82323')
                                )
                            )*/
                        );

                        $border_total_array = array(
                            /*'borders' => array(
                                'top' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                                    'color' => array('rgb' => '000000')
                                )
                            )*/
                        );
                    }

                    //COSTI DIRETTI CODICE
                    $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($bg_color);
                    $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($border_array);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $purchase_components['code']);
                    //PRODUZIONI DESCRIZIONE
                    $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($bg_color);
                    $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($border_array);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($purchase_components['description'], 0, 30));
                    //MATERIALI qta
                    $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($bg_color);
                    $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray($border_array);
                    $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow)->applyFromArray(array(
                            'borders' => array(
                                'right' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                    'color' => array('rgb' => 'cccccc')
                                )
                            )
                    ));
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(5);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    //$qty = floatval($purchase_components['qty']);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow,number_format($purchase_components['qty'],2,',',''));
                    //$this->CI->excel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(2).$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
                    //COSTO UNITARIO
                    $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow. ":E" . $this->xlsRow)->applyFromArray($bg_color);
                    $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow. ":E" . $this->xlsRow)->applyFromArray($border_array);
                    $this->CI->excel->getActiveSheet()->getStyle("D" . $this->xlsRow. ":E" . $this->xlsRow)->applyFromArray(array(
                        'borders' => array(
                            'right' => array(
                                'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                'color' => array('rgb' => 'cccccc')
                            )
                        )
                    ));
                    $this->CI->excel->getActiveSheet()->mergeCells("D" . $this->xlsRow . ":E" . $this->xlsRow);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(3)->setWidth(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(3,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(3,$this->xlsRow, number_format($purchase_components['cost'], 2, ',',''));
                    //MATERIALI TOTALE
                    $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow )->applyFromArray($bg_color);
                    $this->CI->excel->getActiveSheet()->getStyle("F" . $this->xlsRow )->applyFromArray($border_array);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(5)->setWidth(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(5,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyle("F".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $this->CI->excel->getActiveSheet()->setCellValue("F".$this->xlsRow, '=C'.$this->xlsRow . '*D' .$this->xlsRow);

                    //VUOTO
                    $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow . ":J" . $this->xlsRow)->applyFromArray($puchase_white_style);
                    $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow . ":J" . $this->xlsRow)->applyFromArray($border_total_array);
                    //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(20);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->mergeCells("G"  . $this->xlsRow . ":J" .  $this->xlsRow);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow);

                    $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($puchase_white_style);
                    $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($border_total_array);


                    if($old_mid_category != $current_mid_category){
                        $this->xlsRowValue['acquisti']['L'][] = $this->xlsRow;
                        $old_mid_category = $current_mid_category;
                    }

                    $tot_purc_components += $purchase_components['qty'] * $purchase_components['cost'];
                    $counter++;
                    $this->xlsRow++;
                    $lastRow = $this->xlsRow;
                }
            }
            $conta = count($this->xlsRowValue['acquisti']['L']);
            for($i = 0; $i < $conta; $i++) {
                if($i == $conta - 1) {
                    $end = $lastRow-1;
                } else {
                    $end = next($this->xlsRowValue['acquisti']['L'])-1;
                }
                $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRowValue['acquisti']['L'][$i] . ":A" . $end)->applyFromArray($costi_indiretti_code_style);
                $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRowValue['acquisti']['L'][$i] . ":B" . $end)->applyFromArray($costi_indiretti_code_style);
                $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRowValue['acquisti']['L'][$i] . ":F" . $end)->applyFromArray($costi_indiretti_code_style);
                $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRowValue['acquisti']['L'][$i] . ":J" . $end)->applyFromArray($puchase_style_white_style);
                //$this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRowValue['acquisti']['L'][$i] . ":J" . $end)->applyFromArray($border_array);
                $this->CI->excel->getActiveSheet()->getStyle("K" . $firstRow . ":L" . $lastRow)->applyFromArray($costi_indiretti_code_style);
                $this->CI->excel->getActiveSheet()->getStyle("K" . $firstRow . ":L" . $lastRow)->applyFromArray($costi_indiretti_code_style);
                $this->CI->excel->getActiveSheet()->getStyle('L' . $this->xlsRowValue['acquisti']['L'][$i])->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                $this->CI->excel->getActiveSheet()->setCellValue('L' . $this->xlsRowValue['acquisti']['L'][$i], '=SUM(E' . $this->xlsRowValue['acquisti']['L'][$i] . ':F' . $end . ')');
            }

            //riga totale
            if(!empty($product_cost_card_data['purchase_components_cost']['totals'])) {
                $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                $this->CI->excel->getActiveSheet()->getStyle('L'. ($this->xlsRow))->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                $this->CI->excel->getActiveSheet()->setCellValue('L'. ($this->xlsRow),'=SUM(L' . $firstRow . ':L' .($lastRow-1) .')');
                $this->xlsRowValue['riepilogo_acquisti'] = $this->xlsRow;
            }
        }


        $lastRow = (!empty($lastRow) && isset($lastRow)) ? $lastRow : $this->xlsRow;
        $this->CI->excel->getActiveSheet()->getStyle("K" . $firstRow . ":L" . $lastRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($this->total_style);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray(array(
            'borders' => array(
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            )
        ));
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":J" . $this->xlsRow)->applyFromArray($puchase_style_white_style);
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow) . ":J" . ($this->xlsRow))->applyFromArray(array(
                'borders' => array(
                    'top' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => array('rgb' => '000000')
                    )
                )
            )
        );
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow+1) . ":K" . ($this->xlsRow+1))->applyFromArray($this->fill_white);

        $this->xlsRow++;
        $result['tot_purc_components']  = $tot_purc_components;
        return $result;
    }

    public function export_assembly_cost($product_cost_card_data,$standard_td) {
        $this->export_assembly_cost_header($standard_td);

        $res = $this->export_assembly_cost_body($product_cost_card_data,$standard_td);

        $result['tot_asse_components']  = $res['tot_asse_components'];

        return $result;

    }

    public function export_assembly_cost_header($standard_td) {
        $this->xlsRow++;

        //COSTI DIRETTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . (intval($this->xlsRow)+1))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("A" . $this->xlsRow . ":A" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, strtoupper(lang('direct_costs')));
        //ACQUISTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)));
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('assembly')));

        $this->CI->excel->getActiveSheet()->getStyle("C"  . $this->xlsRow . ":F" .  intval(($this->xlsRow+1)))->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("C"  . $this->xlsRow . ":F" .  intval(($this->xlsRow+1)));
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2,5)->setWidth(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

        //TEMPI
        $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow . ":J" . $this->xlsRow)->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(80);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6, $this->xlsRow, strtoupper(lang('times')));
        //TEMPI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("G" . ($this->xlsRow+1))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,($this->xlsRow+1))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,($this->xlsRow+1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6, ($this->xlsRow+1), strtoupper(lang('label_time')));
        //TEMPI QTA HEAD
        $this->CI->excel->getActiveSheet()->getStyle("H" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,($this->xlsRow+1))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,($this->xlsRow+1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7, (intval($this->xlsRow+1)), strtoupper(lang('%_carico_scarico_short')));
        //TEMPI COSTO UNITARIO
        $this->CI->excel->getActiveSheet()->getStyle("I" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,($this->xlsRow+1))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,($this->xlsRow+1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8, (intval($this->xlsRow+1)), strtoupper(lang('%_ineff_short')));
        //TEMPI TOTALE
        $this->CI->excel->getActiveSheet()->getStyle("J" . (intval($this->xlsRow+1)))->applyFromArray($this->standard_td_border_top_white);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,($this->xlsRow+1))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,($this->xlsRow+1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(9, (intval($this->xlsRow+1)), strtoupper(lang('unit_cost_short')));



        //TOTALE HEAD
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($standard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)));
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(10);

        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['totals'], $this->xlsRow, strtoupper(lang('total')));
        $this->xlsRow++;
        $this->xlsRow++;
    }

    public function export_assembly_cost_body($product_cost_card_data,$standard_td) {
        $assemby_cost_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            /*'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $assemby_cost_white_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            /*'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $tot_assembly_cost = 0;
        $html = '';
        if(!empty($product_cost_card_data['assembly_costs'])) {
            $tot_assembly_cost = $product_cost_card_data['assembly_costs']['total'];
                    $startRow= $this->xlsRow;
                    //COSTI DIRETTI CODICE
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, '');
                    //ASSEMBLAGGIO DESCRIZIONE
                    $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, lang('assembly_detail'));

                    //VUOTO
            $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow . ":F" . $this->xlsRow)->applyFromArray($assemby_cost_style);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            // $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->mergeCells("C"  . $this->xlsRow . ":F" .  $this->xlsRow);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

                    //ASSEMBLAGGIO TEMPO
            $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow)->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                        'color' => array('rgb' => 'CCCCCC')
                    )
                )
            ));
            // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow,number_format($product_cost_card_data['assembly_costs']['assembly_time'], 2, ',',''));

            $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                        'color' => array('rgb' => 'CCCCCC')
                    )
                )
            ));
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow,number_format($product_cost_card_data['assembly_costs']['charge_percentage'], 2, ',',''));

                    //COSTO UNITARIO
            $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                        'color' => array('rgb' => 'CCCCCC')
                    )
                )
            ));
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow, '');
                    //MATERIALI TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow )->applyFromArray($assemby_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(9,$this->xlsRow,  number_format($product_cost_card_data['assembly_costs']['department_hour_cost'], 2, ',', ''));

            $this->CI->excel->getActiveSheet()->getStyle("L" . $this->xlsRow )->applyFromArray($assemby_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($assemby_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(10)->setWidth(10);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                    $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                    $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, "=G" . $this->xlsRow ."*60*(1+H" . $this->xlsRow ."/100)*J" . $this->xlsRow . "/3600");
                    $this->xlsRow++;
        }
        /*KITs*/
        $tot_kits_cost      = 0;
        $assembly_dep_perc  = 0;
        $year               = $this->CI->utility->get_field_value('year', $this->CI->db->tb_article_cost, array('id_article_cost' => $product_cost_card_data['id_article_cost']));
        $department_id      = $this->CI->utility->get_field_value('id_department', $this->CI->db->tb_department, array('code' => 'CFZ'));
        if($year && $department_id) {
            $effect_percentage = $this->CI->utility->get_field_value('effect_percentage', $this->CI->db->tb_tab_department_worked, array('year' => $year, 'fk_department_id' => $department_id));
            if ($effect_percentage) {
                $assembly_dep_perc = $effect_percentage;
            }
        }
        if(!empty($product_cost_card_data['production_components_cost'])) {
            if (!empty($product_cost_card_data['production_components_cost']['production_components'])) {
                $counter = 0;
                foreach ($product_cost_card_data['production_components_cost']['production_components'] as $production_component) {
                    if (substr($production_component['code'], 0, 3) == '60.') {
                        $bg_color = '';
                        /*if($counter%2 != 0){
                            $bg_color = 'bgcolor="#eaeaea"';
                        }*/

                        //COSTI DIRETTI CODICE
                        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(0)->setWidth(25);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(0, $this->xlsRow, $production_component['code']);
                        //ASSEMBLAGGIO DESCRIZIONE
                        $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(1)->setWidth(50);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(1, $this->xlsRow, substr($production_component['description'], 0, 30));

                        //VUOTO
                        $this->CI->excel->getActiveSheet()->getStyle("C" . $this->xlsRow . ":F" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2)->setWidth(20);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->mergeCells("C"  . $this->xlsRow . ":F" .  $this->xlsRow);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

                        //ASSEMBLAGGIO TEMPO
                        $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        $this->CI->excel->getActiveSheet()->getStyle("G" . $this->xlsRow)->applyFromArray(array(
                            'borders' => array(
                                'right' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                    'color' => array('rgb' => 'CCCCCC')
                                )
                            )
                        ));
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6)->setWidth(10);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow,number_format($production_component['theoric_time'], 2, ',', ''));

                        $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        $this->CI->excel->getActiveSheet()->getStyle("H" . $this->xlsRow)->applyFromArray(array(
                            'borders' => array(
                                'right' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                    'color' => array('rgb' => 'CCCCCC')
                                )
                            )
                        ));
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(7)->setWidth(10);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(7,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(7,$this->xlsRow,'');

                        //COSTO UNITARIO
                        $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        $this->CI->excel->getActiveSheet()->getStyle("I" . $this->xlsRow)->applyFromArray(array(
                            'borders' => array(
                                'right' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                    'color' => array('rgb' => 'CCCCCC')
                                )
                            )
                        ));
                       // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(8)->setWidth(10);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(8,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(8,$this->xlsRow,  number_format($assembly_dep_perc, 2, ',', ''));
                        //MATERIALI TOTALE
                        $this->CI->excel->getActiveSheet()->getStyle("J" . $this->xlsRow )->applyFromArray($assemby_cost_white_style);
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(9)->setWidth(10);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(9,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(9,$this->xlsRow,  number_format($production_component['hour_cost'], 2, ',', ''));

                        $this->CI->excel->getActiveSheet()->getStyle("L" . $this->xlsRow )->applyFromArray($assemby_cost_white_style);
                        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($assemby_cost_white_style);
                        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(10)->setWidth(20);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
                        $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
                        $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, "=G" . $this->xlsRow ."*60*(1+I" . $this->xlsRow ."/100)*J" . $this->xlsRow . "/3600");

                        $tot_kits_cost += $production_component['production_cost'] * $production_component['qty'];
                        $endRow= $this->xlsRow;
                        $this->xlsRow++;

                        $counter++;
                    }
                }
            }
        }

        $costi_indiretti_code_style = array(
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
        );
        $endRow = (isset($endRow) && $endRow != '' &&  !empty($endRow)) ?  $endRow : $startRow;

        $this->CI->excel->getActiveSheet()->getStyle("A" . $startRow . ":A" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("B" . $startRow . ":B" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("C" . $startRow . ":F" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("G" . $startRow . ":J" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $startRow . ":L" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $startRow . ":L" . $endRow)->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow) . ":J" . ($this->xlsRow))->applyFromArray($this->fill_white);
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow+1) . ":K" . ($this->xlsRow+1))->applyFromArray($this->fill_white);
        //riga totale
        if(isset($product_cost_card_data['assembly_costs']['total'])) {
            //$this->CI->excel->getActiveSheet()->setCellValue('L'. ($this->xlsRow),number_format($tot_kits_cost + $product_cost_card_data['assembly_costs']['total'], 2, ',', '.'));
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($this->total_style);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue('L'. ($this->xlsRow),'=SUM(L' . $startRow . ':L' .($endRow) .')');
            $this->xlsRowValue['riepilogo_assemblaggio'] = $this->xlsRow;
        }

        $result['tot_asse_components']  = $tot_kits_cost + $product_cost_card_data['assembly_costs']['total'];
        //$result['html']                 = $html;

        $this->xlsRow++;
        return $result;
    }

    public function export_indirect_cost($product_cost_card_data,$standard_td) {

        $this->export_indirect_cost_header($product_cost_card_data, $standard_td);

        $res = $this->export_indirect_cost_body($product_cost_card_data, $standard_td);

        $result['tot_indi_components']  = $res['tot_indi_components'];

        return $result;
    }

    public function export_indirect_cost_header($product_cost_card_data, $standard_td) {
        $this->xlsRow++;
        //COSTI DIRETTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . $this->xlsRow)->applyFromArray($standard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("A" . $this->xlsRow . ":A" . $this->xlsRow);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(40);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, strtoupper(lang('indirect_cost')));
        //ACQUISTI HEAD
        $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . $this->xlsRow)->applyFromArray($standard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("B" . $this->xlsRow . ":B" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, '');

        $this->CI->excel->getActiveSheet()->getStyle("C"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($standard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("C"  . $this->xlsRow . ":J" .  $this->xlsRow);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2,5)->setWidth(20);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');


        //TOTALE HEAD
        $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . $this->xlsRow)->applyFromArray($standard_td);
        $this->CI->excel->getActiveSheet()->mergeCells("K" . $this->xlsRow . ":L" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(25);
        // $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getFont()->setBold(true);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['totals'], $this->xlsRow, strtoupper(lang('total')));
        $this->xlsRow++;
        //$this->xlsRow++;

    }

    public function export_indirect_cost_body($product_cost_card_data, $standard_td) {
        $indirect_cost_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            /*'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $indirect_cost_white_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            /*'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),*/
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        if(!empty($product_cost_card_data['indirect_costs'])) {

            //COSTI INDIRETTI RICAVI
            $costi_indiretti_start_row = $this->xlsRow;
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . (intval($this->xlsRow)+1))->applyFromArray($indirect_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, '');

            $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('revenue')) . ' ' . $product_cost_card_data['indirect_costs']['year']);

            $this->CI->excel->getActiveSheet()->getStyle("C"  . $this->xlsRow . ":J" .  intval(($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2,5)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

            //TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => array('rgb' => '000000')
                    )
                ),
            ));
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(11, $this->xlsRow, number_format($product_cost_card_data['indirect_costs']['indirect_revenue'], 2, ',', '.'));
            $this->xlsRow++;


            //COSTI INDIRETTI
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . (intval($this->xlsRow)+1))->applyFromArray($indirect_cost_white_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, '');

            $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_white_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('indirect_cost_detail')) . ' ' . $product_cost_card_data['indirect_costs']['year']);

            $this->CI->excel->getActiveSheet()->getStyle("C"  . $this->xlsRow . ":J" .  intval(($this->xlsRow+1)))->applyFromArray($indirect_cost_white_style);
            // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2,5)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

            //TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_white_style);
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => array('rgb' => '000000')
                    )
                ),
            ));
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(11, $this->xlsRow, $product_cost_card_data['indirect_costs']['indirect_cost_tot']);
            $this->xlsRow++;


            //% COSTI INDIRETTI SU FATTURATO
            $this->CI->excel->getActiveSheet()->getStyle("A" . $this->xlsRow . ":A" . (intval($this->xlsRow)+1))->applyFromArray($indirect_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['direct_costs'])->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(0,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['direct_costs'], $this->xlsRow, '');

            $this->CI->excel->getActiveSheet()->getStyle("B" . $this->xlsRow . ":B" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn($this->xlsColumns['productions'])->setWidth(50);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(1,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow($this->xlsColumns['productions'], $this->xlsRow, strtoupper(lang('effect_long_descr')));

            $this->CI->excel->getActiveSheet()->getStyle("C"  . $this->xlsRow . ":J" .  intval(($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(2,5)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(2,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(2,$this->xlsRow, '');

            //TOTALE
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray($indirect_cost_style);
            $this->CI->excel->getActiveSheet()->getStyle("K" . $this->xlsRow . ":L" . (intval($this->xlsRow+1)))->applyFromArray(array(
                'borders' => array(
                    'right' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => array('rgb' => '000000')
                    )
                ),
            ));
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11, $this->xlsRow)->setWidth(25);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(11, $this->xlsRow,$product_cost_card_data['indirect_costs']['indirect_perc']);
            $this->xlsRow++;
            $costi_indiretti_end_row = $this->xlsRow;
        }

        if(isset($product_cost_card_data['indirect_costs']['total'])) {
            $sale_price         = $this->CI->utility->get_field_value('sale_price', $this->CI->db->tb_article_cost_detail, array('fk_article_cost_id' => $product_cost_card_data['id_article_cost'], 'fk_price_list_id' => $this->CI->session->userdata('tab_article_cost_filter_price_list')));
            $tot_indirect_cost  = $sale_price * $product_cost_card_data['indirect_costs']['indirect_perc']/100;
            /*$html .= '<tr>
                        <td style="border-top: 0.5px solid black; width: 89%!important;"></td>
                        <td bgcolor="#eaeaea" style="font-size:8pt; width: 11%!important;text-align: right;border: 0.5px solid black;"><b>' . number_format($tot_indirect_cost, 2, ',', '.') . '</b></td>
                    </tr>';*/
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            //$this->CI->excel->getActiveSheet()->setCellValue('L'. ($this->xlsRow),number_format($tot_indirect_cost, 2, ',', '.'));

            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue('L'. $this->xlsRow,'=(L' . ($this->xlsRow+7) . '*L'. ($this->xlsRow-1) . ')/100');
            $this->xlsRowValue['riepilogo_indiretti'] = $this->xlsRow;
        }
        $costi_indiretti_code_style = array(
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
        );
        $costi_indiretti_total_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'eaeaea')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            ),
            'font' => array(
                'bold' => true,
            )
        );
        $costi_indiretti_before_total_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'top' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                    'color' => array('rgb' => '000000')
                )
        );
        $this->CI->excel->getActiveSheet()->getStyle("A" . $costi_indiretti_start_row . ":A" . ($costi_indiretti_end_row-1))->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("B" . $costi_indiretti_start_row . ":B" . ($costi_indiretti_end_row-1))->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("C" . $costi_indiretti_start_row . ":J" . ($costi_indiretti_end_row-1))->applyFromArray($costi_indiretti_code_style);
        $this->CI->excel->getActiveSheet()->getStyle("A" . $costi_indiretti_end_row . ":J" . ($costi_indiretti_end_row))->applyFromArray($costi_indiretti_before_total_style);
        $this->CI->excel->getActiveSheet()->getStyle("K" . $costi_indiretti_end_row . ":L" . ($costi_indiretti_end_row))->applyFromArray($costi_indiretti_total_style);
        $this->CI->excel->getActiveSheet()->getStyle("A" . ($this->xlsRow+1) . ":K" . ($this->xlsRow+1))->applyFromArray($this->fill_white);

        $result['tot_indi_components']  = $tot_indirect_cost;
        //$result['html']                 = $html;
        $this->xlsRow++;
        return $result;
    }

    public function export_total_cost($product_cost_card_data, $tot_no_taxes, $tot_indirect_cost, $taxes, $sale_price, $sale_price_orig,$standard_td, $cost_config_perc){
        $total_cost_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );
        $total_cost_white_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )
        );

        $total_cost_red_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                )
            ),
            'font' => array(
                'color' => array('rgb' => 'ff0000')
            )
        );

        $total_cost_white2_style = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => "FFFFFF")
            ),
            'font' => array(
                'color' => array('rgb' => 'FFFFFF')
            )
        );
        $left_a = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $left_a_finale = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $right_a = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $right_a_finale = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->style['f_lightgrey_bg'])
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $left_b = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $left_b = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $right_b = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                /*'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),*/
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => '000000')
            )

        );
        $right_b_red = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFFFF')
            ),
            'borders' => array(
                /*'left' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),*/
                'right' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000')
                ),
                'bottom'=> array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array('rgb' => 'CCCCCC')
                ),
            ),
            'font' => array(
                'color' => array('rgb' => 'ff0000')
            )

        );
        $font_bold = array(
            'font' => array(
                'bold' => true,
            )
        );
        //TITOLO RIEPILOGO
        $this->xlsRow++;
        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($this->standard_td_border_bottom_none);
        $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":L" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,11)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('riepilogo')));
        $this->xlsRow++;
        //COSTI DIRETTI
        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_a);
        $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('direct_costs')));

        $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_a);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(10,11)->setWidth(10);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, '=L' . $this->xlsRowValue['riepilogo_produzione'] . '+L' . $this->xlsRowValue['riepilogo_acquisti'] . '+L' . $this->xlsRowValue['riepilogo_assemblaggio']);
        $this->xlsRowValue['valore_costi_diretti'] = $this->xlsRow;
        $this->xlsRow++;
        //COSTI INDIRETTI
        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_b);
        $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('indirect_cost')));

        $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_b);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(10,11)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, '=L' . $this->xlsRowValue['riepilogo_indiretti']);
        $this->xlsRowValue['valore_costi_indiretti'] = $this->xlsRow;
        $this->xlsRow++;
        //TASSE
        $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_a);
        $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('taxes')));

        $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_a);
        //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
        $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
        // $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, number_format($taxes, 2, ',', '.'));
		$this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, "=MAX(0,(L" . (intval($this->xlsRow)+2) . "-(L" . (intval($this->xlsRow)-2) . "+L" . (intval($this->xlsRow)-1) . "))/100*" . $cost_config_perc . ')');
        $this->xlsRowValue['valore_tasse'] = $this->xlsRow;
        $this->xlsRow++;
        //CALCULATE
        if(isset($product_cost_card_data['total_costs']['total_cost'])) {
            $tot_post_taxes     = $tot_no_taxes + $taxes + $tot_indirect_cost;
            $perc_post_taxes    = 0;
            if($tot_post_taxes > 0){
                $perc_post_taxes    = ($sale_price - $tot_post_taxes) / $tot_post_taxes * 100;
            }
            //COSTO POST TASSE
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_b);
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($font_bold);
            $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('total_taxes')));
            //COSTO POST TASSE VAL
            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_b);
            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($font_bold);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, "=L" . $this->xlsRowValue['valore_costi_diretti'] . "+L" . $this->xlsRowValue['valore_costi_indiretti'] . "+L" . $this->xlsRowValue['valore_tasse']);
            $this->xlsRowValue['costo_post_tasse'] = $this->xlsRow;
            $this->xlsRow++;
            //PREZZO VENDITA
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_a);
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($font_bold);
            $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
           //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('sale_price')));
            //PREZZO VENDITA VAL
            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_a);
            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($font_bold);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, $sale_price);
            $this->xlsRowValue['prezzo_vendita'] = $this->xlsRow;
            $this->xlsRow++;
            //RICARICO POST TASSE
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_b);
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($font_bold);
            $this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
           // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,8)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('charge_percentage_taxes')));

            $this->CI->excel->getActiveSheet()->getStyle("L"  . $this->xlsRow . ":K" .  $this->xlsRow)->applyFromArray($right_b);
            $this->CI->excel->getActiveSheet()->getStyle("L"  . $this->xlsRow . ":K" .  $this->xlsRow)->applyFromArray($font_bold);
            $this->CI->excel->getActiveSheet()->getStyle("M" .  $this->xlsRow)->applyFromArray($total_cost_white2_style);
            //$this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(11)->setWidth(10);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("L".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->getStyle("M".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue('M'.$this->xlsRow, '=(L' . $this->xlsRowValue['prezzo_vendita'] .'-L' . $this->xlsRowValue['costo_post_tasse'] .')/L' . $this->xlsRowValue['costo_post_tasse'] .'*100');
            $this->CI->excel->getActiveSheet()->setCellValue('L'.$this->xlsRow, '=CONCATENATE(ROUND(M' . $this->xlsRow . ',2), "%")');

            //$this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":K" .  $this->xlsRow)->applyFromArray($total_cost_red_style);
            $this->CI->excel->getActiveSheet()->getStyle("N"  . $this->xlsRow . ":N" .  $this->xlsRow)->applyFromArray($total_cost_white2_style);
           // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->CI->excel->getActiveSheet()->getStyle("K".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->getStyle("N".$this->xlsRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $this->CI->excel->getActiveSheet()->setCellValue("N".$this->xlsRow, "=L" . $this->xlsRowValue['prezzo_vendita']  . "-L" . $this->xlsRowValue['costo_post_tasse']);
            $this->CI->excel->getActiveSheet()->setCellValue('K'.$this->xlsRow, '=CONCATENATE("€", ROUND(N' . $this->xlsRow . ',2))');
            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_b_red);
            $this->xlsRow++;
            //PREZZO DI VENDITA ORIGINALE
            $this->CI->excel->getActiveSheet()->getStyle("G"  . $this->xlsRow . ":J" .  $this->xlsRow)->applyFromArray($left_a_finale);
          // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(6,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            //$this->CI->excel->getActiveSheet()->mergeCells("G" . $this->xlsRow . ":J" . $this->xlsRow);
            $this->CI->excel->getActiveSheet()->setCellValueByColumnAndRow(6,$this->xlsRow, strtoupper(lang('sale_price_orig_short')));

            $this->CI->excel->getActiveSheet()->getStyle("K"  . $this->xlsRow . ":L" .  $this->xlsRow)->applyFromArray($right_a_finale);
           // $this->CI->excel->getActiveSheet()->getColumnDimensionByColumn(6,9)->setWidth(20);
            $this->CI->excel->getActiveSheet()->getRowDimension($this->xlsRow)->setRowHeight(20);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(10,$this->xlsRow)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $this->CI->excel->getActiveSheet()->getStyleByColumnAndRow(11,$this->xlsRow)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $this->CI->excel->getActiveSheet()->setCellValue("L".$this->xlsRow, number_format($sale_price_orig, 2 ,',',''));


        }
    }

    public function sort_production_components_cost($prod_comps){
        $prod_counter = count($prod_comps);
        for($j = 0; $j < $prod_counter; $j++) {
            if($prod_comps[$j]['details']){
                $counter        = count($prod_comps[$j]['details']);
                $counter_acq    = 0;
                for($i = 0; $i < $counter; $i++){
                    if($prod_comps[$j]['details'][$i]['is_purchase']){
                        $temp = $prod_comps[$j]['details'][$counter_acq];
                        $prod_comps[$j]['details'][$counter_acq] = $prod_comps[$j]['details'][$i];
                        $prod_comps[$j]['details'][$i] = $temp;
                        $counter_acq++;
                    }
                }
            }
        }
        return $prod_comps;
    }

    public function sort_production_subcomponents_cost($prod_comps){
        $counter        = count($prod_comps);
        $counter_acq    = 0;
        for($i = 0; $i < $counter; $i++){
            if($prod_comps[$i]['is_purchase']){
                $temp = $prod_comps[$counter_acq];
                $prod_comps[$counter_acq] = $prod_comps[$i];
                $prod_comps[$i] = $temp;
                $counter_acq++;
            }
        }
        return $prod_comps;
    }

}
