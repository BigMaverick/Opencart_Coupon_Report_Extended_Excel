<?php


class ControllerReportSaleCoupon extends Controller {
	public function index() {
		$this->load->language('report/sale_coupon');

		$this->document->setTitle($this->language->get('heading_title'));

		if (isset($this->request->get['filter_date_start'])) {
			$filter_date_start = $this->request->get['filter_date_start'];
		} else {
			$filter_date_start = '';
		}

		if (isset($this->request->get['filter_date_end'])) {
			$filter_date_end = $this->request->get['filter_date_end'];
		} else {
			$filter_date_end = '';
		}

		if (isset($this->request->get['page'])) {
			$page = $this->request->get['page'];
		} else {
			$page = 1;
		}

		$url = '';

		if (isset($this->request->get['filter_date_start'])) {
			$url .= '&filter_date_start=' . $this->request->get['filter_date_start'];
		}

		if (isset($this->request->get['filter_date_end'])) {
			$url .= '&filter_date_end=' . $this->request->get['filter_date_end'];
		}

		if (isset($this->request->get['page'])) {
			$url .= '&page=' . $this->request->get['page'];
		}

		$data['breadcrumbs'] = array();

		$data['breadcrumbs'][] = array(
			'text' => $this->language->get('text_home'),
			'href' => $this->url->link('common/dashboard', 'token=' . $this->session->data['token'], 'SSL')
		);

		$data['breadcrumbs'][] = array(
			'text' => $this->language->get('heading_title'),
			'href' => $this->url->link('report/sale_coupon', 'token=' . $this->session->data['token'] . $url, 'SSL')
		);

		$d= $this->load->model('report/coupon');
		
		$data['coupons'] = array();

		$filter_data = array(
			'filter_date_start'	=> $filter_date_start,
			'filter_date_end'	=> $filter_date_end,
			'start'             => ($page - 1) * $this->config->get('config_limit_admin'),
			'limit'             => $this->config->get('config_limit_admin')
		);

		

		$coupon_total = $this->model_report_coupon->getTotalModifiedCoupons($filter_data);


		$results = $this->model_report_coupon->getModifiedCoupons($filter_data);

		foreach ($results as $result) {
			$data['coupons'][] = array(
				'order_id'   => $result['order_id'],
				'name'   => $result['name'],
				'model' => $result['model'],
				'original_price'   => $result['original_price'],
				'code'   => $result['code'],
				'savings' => $result['savings'],
				'sale_price'   => $result['sale_price'],
				'firstname' => $result['firstname'],
				'date_added' => $result['date_added'],
				'payment_city' => $result['payment_city'],
				'payment_state' => $result['payment_state'],
				'edit'   => $this->url->link('marketing/coupon/edit', 'token=' . $this->session->data['token'] . '&coupon_id=' . $result['coupon_id'] . $url, 'SSL')
			);
		}

		$data['heading_title'] = $this->language->get('heading_title');

		$data['text_list'] = $this->language->get('text_list');
		$data['text_no_results'] = $this->language->get('text_no_results');
		$data['text_confirm'] = $this->language->get('text_confirm');

		$data['column_name'] = $this->language->get('column_name');
		$data['column_code'] = $this->language->get('column_code');
		$data['column_order_id'] = $this->language->get('column_order_id');
		$data['column_model'] = $this->language->get('column_model');
		$data['column_date_added'] = $this->language->get('column_date_added');
		$data['column_original_price'] = $this->language->get('column_original_price');
		$data['column_savings'] = $this->language->get('column_savings');
		$data['column_sale_price'] = $this->language->get('column_sale_price');
		$data['column_firstname'] = $this->language->get('column_firstname');
		$data['column_payment_city'] = $this->language->get('column_payment_city');
		$data['column_payment_state'] = $this->language->get('column_payment_state');


		$data['entry_date_start'] = $this->language->get('entry_date_start');
		$data['entry_date_end'] = $this->language->get('entry_date_end');

		$data['button_edit'] = $this->language->get('button_edit');
		$data['button_filter'] = $this->language->get('button_filter');

		$data['token'] = $this->session->data['token'];

		$url = '';

		if (isset($this->request->get['filter_date_start'])) {
			$url .= '&filter_date_start=' . $this->request->get['filter_date_start'];
		}

		if (isset($this->request->get['filter_date_end'])) {
			$url .= '&filter_date_end=' . $this->request->get['filter_date_end'];
		}

		$pagination = new Pagination();
		$pagination->total = $coupon_total;
		$pagination->page = $page;
		$pagination->limit = $this->config->get('config_limit_admin');
		$pagination->url = $this->url->link('report/sale_coupon', 'token=' . $this->session->data['token'] . $url . '&page={page}', 'SSL');

		$data['pagination'] = $pagination->render();

		$data['results'] = sprintf($this->language->get('text_pagination'), ($coupon_total) ? (($page - 1) * $this->config->get('config_limit_admin')) + 1 : 0, ((($page - 1) * $this->config->get('config_limit_admin')) > ($coupon_total - $this->config->get('config_limit_admin'))) ? $coupon_total : ((($page - 1) * $this->config->get('config_limit_admin')) + $this->config->get('config_limit_admin')), $coupon_total, ceil($coupon_total / $this->config->get('config_limit_admin')));

		$data['filter_date_start'] = $filter_date_start;
		$data['filter_date_end'] = $filter_date_end;

		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');

		$this->response->setOutput($this->load->view('report/sale_coupon.tpl', $data));
	}

	public function excel () {
		require_once DIR_SYSTEM . 'library/excel/PHPExcel.php';
		require_once DIR_SYSTEM . 'library/excel/PHPExcel/IOFactory.php';
		 $this->load->model('report/coupon');

		 	


		if (isset($this->request->get['filter_date_start'])) {
			$filter_date_start = $this->request->get['filter_date_start'];
		} else {
			$filter_date_start = '';
		}

		if (isset($this->request->get['filter_date_end'])) {
			$filter_date_end = $this->request->get['filter_date_end'];
		} else {
			$filter_date_end = '';
		}


		 $filter_data = array(
			'filter_date_start'	=> $filter_date_start,
			'filter_date_end'	=> $filter_date_end,
		);
		 
		// Create new PHPExcel object
		$objPHPExcel = new PHPExcel();
// Set document properties
		$objPHPExcel->getProperties()->setCreator("Comcast")
		->setLastModifiedBy("Comcast")
		->setTitle("Office 2007 XLSX Test Document")
		->setSubject("Office 2007 XLSX Test Document")
		->setDescription("Coupon Reports for Comcast.")
		->setKeywords("office 2007 openxml php")
		->setCategory("Comcast");




		$results = $this->model_report_coupon->getModifiedCoupons($filter_data);

		if(!sizeof($results)) {
			// echo "saldfjk";
			$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A1', 'No Data available');

			
		} else {


			// Add Column data
		$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue('A1', 'OrderID')
		->setCellValue('B1', 'Product Name')
		->setCellValue('C1', 'Product Model')
		->setCellValue('D1', 'Date')
		->setCellValue('E1', 'Original Price')
		->setCellValue('F1', 'Coupon Code')
		->setCellValue('G1', 'Savings')
		->setCellValue('H1', 'Sale Price')
		->setCellValue('I1', 'Name')
		->setCellValue('J1', 'City')
		->setCellValue('K1', 'State');



			foreach ($results as $key => $result) {
			# code...
				$key += 2;
				$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue('A'.$key, $result['order_id'])
				->setCellValue('B'.$key, $result['name'])
				->setCellValue('C'.$key, $result['model'])
				->setCellValue('D'.$key, $result['date_added'])
				->setCellValue('E'.$key, $result['original_price'])
				->setCellValue('F'.$key, $result['code'])
				->setCellValue('G'.$key, $result['savings'])
				->setCellValue('H'.$key, $result['sale_price'])
				->setCellValue('I'.$key, $result['firstname'])
				->setCellValue('J'.$key, $result['payment_city'])
				->setCellValue('K'.$key, $result['payment_state']);
			}

			$sheet = $objPHPExcel->getActiveSheet();
			$lastrow = $objPHPExcel->getActiveSheet()->getHighestRow();

			$objPHPExcel->getActiveSheet()
			->getStyle('F1:F'.$lastrow)
			->getAlignment()
			->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

			$objPHPExcel->getActiveSheet()
			->getStyle('A1:A'.$lastrow)
			->getAlignment()
			->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);


			$objPHPExcel->getActiveSheet()->getStyle("A1:K1")->applyFromArray(array("font" => array( "bold" => true)));
			$sheet->getColumnDimension()->setAutoSize(true);

			for($col = 'A'; $col !== 'L'; $col++) {
				$objPHPExcel->getActiveSheet()
				->getColumnDimension($col)
				->setAutoSize(true);
			}
		}

// Rename worksheet
		$objPHPExcel->getActiveSheet()->setTitle('Comcast Coupon Report');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);
// Redirect output to a client’s web browser (Excel5)
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="Comcast_Coupon_Report.xls"');
		header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
	-	header('Cache-Control: max-age=1');
// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;
	}

	

}