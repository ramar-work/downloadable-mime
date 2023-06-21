<?php
// open lib
include_once( "xlsxwriter.class.php" );

// can't use these on the web, but I can use them now
ini_set('display_errors', 0);
ini_set('log_errors', 1);
error_reporting(E_ALL & ~E_NOTICE);

// get instance of writer
$w = new XLSXWriter();
$w->setAuthor( "NCAT report writer app" );

// define a place for the report
$filename = "report.xlsx";

// header format
$header_format = array(
	"font-size" => 18,
	"font-style" => "bold",
	"height" => 30,
	"fill" => "#eee",
	"color" => "#000",
	"border" => "bottom",
);

$val_format = array(
	"font-size" => 12,
	"color" => "#000",
	"valign" => "center",
	"wrap_text" => true,
	"height" => 50
);

// summary format
$summary_format = array(
	"font-size" => 14,
	"color" => "#000",
	"valign" => "center",
	"height" => 20,
	"wrap_text" => true,
);


// load json from whatever
$src = file_get_contents( __DIR__ . "/dored.json" );
$obj = json_decode( $src );
//print_r( $obj );
//print_r( $obj->team );

// what data do I want?
$srows = array(
	array( "Team", $obj->team ),
	array( "Team Leader", $obj->team_leader ),
	array( "External Analysis", $obj->analysis_external ),
	array( "Internal Analysis", $obj->analysis_internal ),
	array( "Team Members", implode( ", ", $obj->team_members ) ),
);

// create a header for the summary sheet
$w->writeSheetHeader('Summary', 
	array('Keys' => "string" ,'Values' => "string"), $col_options = ['widths'=>[30,120]] );

$w->writeSheetRow( 'Summary', array( "Plan Summary" ), $header_format );
$w->writeSheetRow( 'Summary', array( "" ) );

// create and try writing to the summary sheet 
foreach( $srows as $row ) {
	$w->writeSheetRow( 'Summary', $row, $summary_format );
}

// each goal can get a sheet of its own as well
$goals = $obj->goals;

// must be able to reference the goals
$goal_names = array(
	"Transformative Engagement"
, "Leadership and Innovation"
, "Performance Excellence"
, "Collaborative and Inclusive Culture"
, "Responsive Scholarship and Impact"
);

// move through each of these and add them 
$gcount = 1;
foreach ( $goals as $gouter ) {
	// Not written so well... and don't know the goal names...	
	$g = $gouter[0];

	// Define names
	$name = "Goal $gcount";
	$gindex = $gcount - 1;
	$title = "Goal $gcount - ${goal_names[$gindex]}";

	$w->writeSheetHeader($name, 
		array('Keys' => "string" ,'Values' => "string"), $col_options = ['widths'=>[30,150]] );

	// Start with writing the as the first row and a blank row after
	$w->writeSheetRow( $name, array( $title ), $header_format );	
	$w->writeSheetRow( $name, array( '' ) );	

	// Strategy
	$w->writeSheetRow( $name, array( "Strategy", $g->strategy ), $val_format );	

	// Action 
	$w->writeSheetRow( $name, array( "Action", $g->action ), $val_format );	

	// Measure
	$w->writeSheetRow( $name, array( "Measure", $g->measure ), $val_format );	

	// Metrics 
	$w->writeSheetRow( $name, array( "Metrics" ), $val_format );	 
	$mcount = 1;
	foreach ( $g->metrics as $m ) {
		$w->writeSheetRow( $name, array( "${mcount}.", $m ), $val_format );	
		$mcount++;
	}

	// Baseline	
	$w->writeSheetRow( $name, array( "Baseline", $g->baseline ), $val_format );	 

	// Target 
	$w->writeSheetRow( $name, array( "Target", $g->target ), $val_format );	 

	// Increment
	$gcount++;
}

// wait on the metrics, this is still very unclear 
// ...

// write the file
// TODO: Use a random name
$w->writeToFile( "temp.xlsx" );


// Send it as an HTTP message if that's the case...
if ( array_key_exists( "REQUEST_METHOD", $_SERVER ) ) {

	// Read into stack memory	
	$mem = file_get_contents( "temp.xlsx" );
	
	// Wow, what a gnarly mimetype
	$mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	// Define a report name
	$report_name = "\"report.xlsx\"";

	// Send a response
	http_response_code( 200 );
	header( "Content-Type: $mime" );
	header( "Content-Length: " . strlen( $mem ) ); 
	header( "Content-Disposition: attachment; filename=$report_name" );
	//file_put_contents( "php://stdout", $mem  );
	print( $mem );
}


// get rid of file
unlink( "temp.xlsx" );
?>
