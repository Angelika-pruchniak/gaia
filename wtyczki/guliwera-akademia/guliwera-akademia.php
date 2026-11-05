<?php
/**
 * Plugin Name: Guliwera – Akademia (aplikacja)
 * Description: Udostępnia aplikację Akademia Guliwera pod adresem guliwera4.pl/akademia/ (aplikacja na telefon do zainstalowania z przeglądarki). Zostaw włączoną.
 * Version: 1.0
 */
if ( ! defined( 'ABSPATH' ) ) { exit; }

define( 'GK_AK_PREFIX', '/akademia/' );

/**
 * Plik pod /akademia/... wysyłamy prosto z folderu app/ wtyczki,
 * zanim WordPress zacznie budować stronę.
 */
function gk_ak_serve() {
	$path = wp_parse_url( isset( $_SERVER['REQUEST_URI'] ) ? wp_unslash( $_SERVER['REQUEST_URI'] ) : '', PHP_URL_PATH );
	if ( ! is_string( $path ) ) { return; }
	if ( '/akademia' === $path ) {
		wp_redirect( home_url( GK_AK_PREFIX ), 301 );
		exit;
	}
	if ( 0 !== strpos( $path, GK_AK_PREFIX ) ) { return; }

	$rel = rawurldecode( substr( $path, strlen( GK_AK_PREFIX ) ) );
	if ( '' === $rel ) { $rel = 'index.html'; }

	$base = realpath( __DIR__ . '/app' );
	$file = ( false !== strpos( $rel, '..' ) || ! preg_match( '#^[A-Za-z0-9._/-]+$#', $rel ) )
		? false
		: realpath( $base . '/' . $rel );
	if ( ! $file || 0 !== strpos( $file, $base . DIRECTORY_SEPARATOR ) || ! is_file( $file ) ) {
		// Nieznany adres w aplikacji -> ekran główny aplikacji.
		$rel  = 'index.html';
		$file = $base . '/index.html';
	}

	$types = array(
		'html'        => 'text/html; charset=utf-8',
		'js'          => 'text/javascript; charset=utf-8',
		'css'         => 'text/css; charset=utf-8',
		'svg'         => 'image/svg+xml',
		'png'         => 'image/png',
		'json'        => 'application/json',
		'webmanifest' => 'application/manifest+json',
	);
	$ext = strtolower( pathinfo( $file, PATHINFO_EXTENSION ) );

	// Nie cache'ujemy przez wtyczki stron (W3 Total Cache, WP-Optimize) i nie minifikujemy.
	foreach ( array( 'DONOTCACHEPAGE', 'DONOTMINIFY', 'DONOTCDN', 'DONOTCACHEOBJECT' ) as $c ) {
		if ( ! defined( $c ) ) { define( $c, true ); }
	}
	while ( ob_get_level() > 0 ) { ob_end_clean(); }

	http_response_code( 200 );
	header( 'Content-Type: ' . ( isset( $types[ $ext ] ) ? $types[ $ext ] : 'application/octet-stream' ) );
	header( 'X-Content-Type-Options: nosniff' );
	if ( 0 === strpos( $rel, 'assets/' ) ) {
		header( 'Cache-Control: public, max-age=31536000, immutable' );
	} else {
		header( 'Cache-Control: no-cache' );
	}
	if ( 'sw.js' === $rel ) {
		header( 'Service-Worker-Allowed: ' . GK_AK_PREFIX );
	}
	header( 'Content-Length: ' . filesize( $file ) );
	readfile( $file );
	exit;
}
add_action( 'plugins_loaded', 'gk_ak_serve', 0 );
