<?php

use \Mantis\Export\TableExportProvider;

class MantisLegacyExportPlugin extends \MantisPlugin {
	function register() {
		$this->name = 'MantisLegacyExport';
		$this->description = 'MantisLegacyExport';
		//$this->page = "config_page";

		$this->version = MANTIS_VERSION;
		$this->requires = array(
			'MantisCore' => '2.22.0-dev',
		);

		$this->author = 'MantisBT Team';
		$this->contact = 'mantisbt-dev@lists.sourceforge.net';
		$this->url = 'http://www.mantisbt.org';
	}

	function init() {
		require_once( plugin_path() . 'include' . DIRECTORY_SEPARATOR . 'csv_api.php' );
		require_once( plugin_path() . 'include' . DIRECTORY_SEPARATOR . 'excel_api.php' );
	}

	function hooks() {
		$t_hooks = array();
		$t_hooks['EVENT_EXPORT_DISCOVERY'] = 'ev_export_discovery';
		$t_hooks['EVENT_EXPORT_REQUEST'] = 'ev_export_request';
		return $t_hooks;
	}

	function ev_export_discovery( $p_event ) {
		$t_csv = new TableExportProvider();
		$t_csv->unique_id = 'MantisLegacyExport_csv';
		$t_csv->file_extension = 'csv';
		$t_csv->short_name = 'Text CSV';
		$t_csv->provider_name = 'Mantis';

		$t_excel = new TableExportProvider();
		$t_excel->unique_id = 'MantisLegacyExport_excel';
		$t_excel->file_extension = 'xml';
		$t_excel->short_name = 'Excel XML';
		$t_excel->provider_name = 'Mantis';

		return array( $t_csv, $t_excel );
	}

	function ev_export_request( $p_event, $p_id ) {
		switch( $p_id ) {
			case 'MantisLegacyExport_csv':
				require_once( plugin_path() . 'classes' . DIRECTORY_SEPARATOR . 'MantisCsvWriter.php' );
				return new MantisLegacyExport\MantisCsvWriter();
			case 'MantisLegacyExport_excel':
				require_once( plugin_path() . 'classes' . DIRECTORY_SEPARATOR . 'MantisExcelWriter.php' );
				return new MantisLegacyExport\MantisExcelWriter();
		}
	}
}