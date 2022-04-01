// Export Products as columns
// Format - XLS
// Checked - Output column titles as first line
// Button - Export w/o Progressbar
class Woe_Product_Columns_XLS {
    function __construct() {
        //add settings, , skip products        
        add_action("woe_settings_above_buttons", array($this,"draw_options") );
        add_filter("woe_settings_validate_defaults",array($this,"skip_products"),10,1);
    }

    // 1
    function draw_options($settings){
        $selected = !empty($settings[ 'products_as_columns' ]) ? 'checked': '';
        echo '<br><br>
        <input type=hidden name="settings[products_as_columns]" value="0">
        <input type=checkbox name="settings[products_as_columns]" value="1" '. $selected .'>
        <span class="wc-oe-header">Export products as columns,  print <select name="settings[products_as_columns_output_field]" style="width: 100px">
			<option value="qty">Qty</option>
			<option value="line_total">Amount</option>
        </select>
         in cell</span><br>
         Format <b>XLS</b>, button <b>Export w/o progressbar</b>
        <br><br>';
    }

    function skip_products($current_job_settings) {
        if( !empty($current_job_settings['products_as_columns']) )  {
             $current_job_settings["order_fields"]["products"]["checked"] = 0;//  just  skip standard products
             $this->output_field = $current_job_settings['products_as_columns_output_field'];
             // read orders
             add_action("woe_order_export_started",array($this,"start_new_order"),10,1);

             //stop default output for rows
             add_action("woe_xls_header_filter",array($this,"prepare_xls_vars"),10,2);
             add_action("woe_xls_output_filter",array($this,"record_xls_rows"),10,2);
             add_action("woe_xls_print_footer",array($this,"analyze_products_add_columns"),10,2);
        }
        return $current_job_settings;
    }   

    // 2
    function prepare_xls_vars($data) {
	$this->headers_added = count($data);
        $this->product_columns = array();
        return $data;
    }

    //3
    function start_new_order($order_id) {
        $this->order_id = $order_id;
        return $order_id;
    }
    function record_xls_rows($data,$obj) {
        $order = new WC_Order($this->order_id);
        $extra_cells = array_fill(0, count($this->product_columns), "");
        // work with products
        foreach($order->get_items('line_item') as $item_id=>$item) {
            $product_name = $item['name'];  
			$terms  = get_the_terms( $item->get_product_id(), 'product_tag' );
			if ( $terms ) {
				$arr = array();
				foreach ( $terms as $term )
					$arr[] = $term->name;
				$product_name = join( ",", $arr );
            }
            $pos = array_search($product_name,$this->product_columns);
            if( $pos === false) { // new product detected
                $extra_cells[] = $item[ $this->output_field  ]; 
                $this->product_columns[] = $product_name;
            } else {
                $extra_cells[$pos] = $item[ $this->output_field  ]; 
            }
        }
        foreach($extra_cells as $pc)
            $data[] = $pc;
        return $data;
    }

    //4 
    function analyze_products_add_columns($phpExcel,$formatter) {
        // add products as titles
	    foreach($this->product_columns as $pos=>$text) {
            	$formatter->objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $pos+$this->headers_added, 1, $text );
	    }
        //make first bold
        $last_column = $formatter->objPHPExcel->getActiveSheet()->getHighestDataColumn();
        $formatter->objPHPExcel->getActiveSheet()->getStyle( "A1:" . $last_column . "1" )->getFont()->setBold( true );
    }
}
new Woe_Product_Columns_XLS();
