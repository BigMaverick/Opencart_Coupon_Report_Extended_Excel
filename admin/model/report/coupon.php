<?php
class ModelReportCoupon extends Model {
	
	public function getModifiedCoupons($data = array()) {
		$sql = "SELECT ch.coupon_id,ch.order_id, oop.name, oop.model, oo.date_added, FORMAT(st.original_price, 2) original_price, oc.code, FORMAT(t.savings, 2) savings, FORMAT(( st.original_price + t.savings ), 2) sale_price, CONCAT(oo.firstname , ' ' ,oo.lastname) firstname, oo.payment_city,oo.payment_zone payment_state from " . DB_PREFIX . "coupon_history ch join " . DB_PREFIX . "coupon oc on ( oc.coupon_id = ch.coupon_id ) join " . DB_PREFIX . "order oo on ( oo.order_id = ch.order_id ) join " . DB_PREFIX . "order_product oop on (oop.order_id = ch.order_id ) join " . DB_PREFIX . "product op on ( oop.product_id = op.product_id ) join ( SELECT order_id, value original_price FROM `" . DB_PREFIX . "order_total` where code = 'sub_total' ) st on (st.order_id = oo.order_id ) join ( SELECT order_id, value savings FROM `" . DB_PREFIX . "order_total` where code = 'coupon' ) t on (t.order_id = oo.order_id )";

		$implode = array();

		if (!empty($data['filter_date_start'])) {
			$implode[] = "DATE(ch.date_added) >= '" . $this->db->escape($data['filter_date_start']) . "'";
		}

		if (!empty($data['filter_date_end'])) {
			$implode[] = "DATE(ch.date_added) <= '" . $this->db->escape($data['filter_date_end']) . "'";
		}

		if ($implode) {
			$sql .= " WHERE " . implode(" AND ", $implode);
		}

		//$sql .= " GROUP BY ch.coupon_id ORDER BY total DESC";

		if (isset($data['start']) || isset($data['limit'])) {
			if ($data['start'] < 0) {
				$data['start'] = 0;
			}

			if ($data['limit'] < 1) {
				$data['limit'] = 20;
			}

			$sql .= " LIMIT " . (int)$data['start'] . "," . (int)$data['limit'];
		}

		$query = $this->db->query($sql);

		return $query->rows;
	}


	public function getTotalModifiedCoupons($data = array()) {
		$sql = "SELECT COUNT(coupon_id) AS total FROM `" . DB_PREFIX . "coupon_history`";

		$implode = array();

		if (!empty($data['filter_date_start'])) {
			$implode[] = "DATE(date_added) >= '" . $this->db->escape($data['filter_date_start']) . "'";
		}

		if (!empty($data['filter_date_end'])) {
			$implode[] = "DATE(date_added) <= '" . $this->db->escape($data['filter_date_end']) . "'";
		}

		if ($implode) {
			$sql .= " WHERE " . implode(" AND ", $implode);
		}

		$query = $this->db->query($sql);

		return $query->row['total'];
	}
}