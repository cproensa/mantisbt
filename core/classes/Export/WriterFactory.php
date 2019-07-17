<?php

namespace Mantis\Export;

class WriterFactory {
	public static function createFromType( $p_type ) {
		switch( $p_type ) {
			case 'csv':
				return new MantisCsvWriter();
		}
	}
	
}
