<?php

use CodeIgniter\Router\RouteCollection;

/**
 * @var RouteCollection $routes
 */
$routes->get('/', static function () {
    return redirect()->to('/public/index.html');
});

$routes->group('api', static function ($routes) {
    $routes->get('dashboard', 'Api::dashboard');
    $routes->get('leaves', 'Api::leaves');
    $routes->post('leaves', 'Api::leaves');
    $routes->get('reports/summary', 'Api::reportsSummary');
    $routes->post('import/excel', 'Api::importExcel');
    $routes->get('users', 'Api::users');
    $routes->post('users', 'Api::users');
    $routes->put('users', 'Api::users');
    $routes->delete('users', 'Api::users');
});
