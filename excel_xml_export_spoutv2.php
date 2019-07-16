<?php
# MantisBT - A PHP based bugtracking system

# MantisBT is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# MantisBT is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with MantisBT.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Excel (2003 SP2 and above) export page
 *
 * @package MantisBT
 * @copyright Copyright 2000 - 2002  Kenzaburo Ito - kenito@300baud.org
 * @copyright Copyright 2002  MantisBT Team - mantisbt-dev@lists.sourceforge.net
 * @link http://www.mantisbt.org
 *
 * @uses core.php
 * @uses authentication_api.php
 * @uses bug_api.php
 * @uses columns_api.php
 * @uses config_api.php
 * @uses excel_api.php
 * @uses file_api.php
 * @uses filter_api.php
 * @uses gpc_api.php
 * @uses helper_api.php
 * @uses print_api.php
 * @uses utility_api.php
 */

# Prevent output of HTML in the content if errors occur
//define( 'DISABLE_INLINE_ERROR_REPORTING', true );

require_once( 'core.php' );
require_api( 'authentication_api.php' );
require_api( 'columns_api.php' );
require_api( 'constant_inc.php' );
require_api( 'csv_api.php' );
require_api( 'file_api.php' );
require_api( 'filter_api.php' );
require_api( 'helper_api.php' );
require_api( 'print_api.php' );

//require_once dirname(__FILE__) . '/library/PHPExcel/Classes/PHPExcel.php';
//require_once 'library/spout/src/Spout/Autoloader/autoload.php';

use Box\Spout\Writer\WriterFactory;
//use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Type;

auth_ensure_user_authenticated();
$time_start = microtime(true);
log_event( LOG_WEBSERVICE, 'SPOUT_XLSX start');

$f_export = gpc_get_string( 'export', '' );

helper_begin_long_process();

# Get current filter
$t_filter = filter_get_bug_rows_filter();

$t_filter_query = new BugFilterQuery( $t_filter );
$t_filter_query->set_limit( EXPORT_BLOCK_SIZE );


# Get the total number of bugs that meet the criteria.
if( 0 == $t_filter_query->get_bug_count() ) {
	print_header_redirect( 'view_all_set.php?type=0&print=1' );
}


# Get columns to be exported
$t_columns = csv_get_columns();

// Create new PHPExcel object
//$objPHPExcel = new PHPExcel();



/*
// Redirect output to a client’s web browser (OpenDocument)
//header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
*/

$writer = WriterFactory::create(Type::XLSX);
//$writer = WriterEntityFactory::createXLSXWriter();
$writer->openToBrowser('export.xlsx');

# export the titles
$t_first_column = true;

$t_titles = array();
foreach ( $t_columns as $t_column ) {
	$t_titles[] = column_get_title( $t_column );
}


//$rowFromValues = WriterEntityFactory::createRowFromArray($t_titles);
//$writer->addRow($rowFromValues);
$writer->addRow($t_titles);

$t_next_row=2;


$block=0;
$t_end_of_results = false;
$t_offset = 0;
//$t_xsheet = $objPHPExcel->getActiveSheet();
do {
	# Clear cache for next block
	bug_clear_cache_all();
	
	# select a new block
	$t_filter_query->set_offset( $t_offset );
	$t_result = $t_filter_query->execute();
	$t_offset += EXPORT_BLOCK_SIZE;

	# Keep reading until reaching max block size or end of result set
	$t_read_rows = array();
	$t_count = 0;
	$t_bug_id_array = array();
	$t_unique_user_ids = array();
	while( $t_count < EXPORT_BLOCK_SIZE ) {
		$t_row = db_fetch_array( $t_result );
		if( false === $t_row ) {
			$t_end_of_results = true;
			break;
		}
		$t_bug_id_array[] = (int)$t_row['id'];
		$t_read_rows[] = $t_row;
		$t_count++;
	}
	# Max block size has been reached, or no more rows left to complete the block.
	# Either way, process what we have

	# convert and cache data
	$t_rows = filter_cache_result( $t_read_rows, $t_bug_id_array );
	bug_cache_columns_data( $t_rows, $t_columns );

	# Clear arrays that are not needed
	unset( $t_read_rows );
	unset( $t_unique_user_ids );
	unset( $t_bug_id_array );

	$t_excel_rows = array();
	# export the rows
	foreach ( $t_rows as $t_row ) {
		$t_first_column = true;

		$t_excel_row = array();
		
		foreach ( $t_columns as $t_column ) {
			$t_value = null;
			if( column_get_custom_field_name( $t_column ) !== null || column_is_plugin_column( $t_column ) ) {
				ob_start();
				$t_column_value_function = 'print_column_value';
				helper_call_custom_function( $t_column_value_function, array( $t_column, $t_row, COLUMNS_TARGET_CSV_PAGE ) );
				$t_value = ob_get_clean();

				//echo csv_escape_string( $t_value );
			} else {
				ob_start();
				$t_function = 'csv_format_' . $t_column;

				echo $t_function( $t_row );
				$t_value = ob_get_clean();
			}
			$t_excel_row[] = $t_value;
			
		}
		//$rowFromValues = WriterEntityFactory::createRowFromArray($t_excel_row);
		//$writer->addRow($rowFromValues);
		$writer->addRow($t_excel_row);
	}
	//$t_xsheet->fromArray($t_excel_rows, null, 'A' . $t_next_row);
	//$t_next_row += EXPORT_BLOCK_SIZE;
$block++;
log_event( LOG_WEBSERVICE, 'SPOUT_XLSX running ('.$block.') mem: ' . memory_get_peak_usage(true));
} while ( false === $t_end_of_results );


$time_end = microtime(true);
$time = $time_end - $time_start;
$mem = memory_get_peak_usage(true);
log_event( LOG_WEBSERVICE, 'SPOUT_XLSX end time: ' . $time);
log_event( LOG_WEBSERVICE, 'SPOUT_XLSX end mem: ' . $mem);


$writer->close();
exit;

