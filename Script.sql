
CREATE TABLE login(
 login_id varchar(8) primary key,
 f_name varchar(14) not null,
 l_name varchar(14) not null,
 email varchar(30) not null,
 password varchar(16) not null,
 c_password varchar(16) not null,
 phone_no number(10) not null
 );
/
CREATE TABLE product_detail
 (
  product_id varchar(10) PRIMARY KEY,
  product_name varchar(30) NOT NULL,
  gst number(5) NOT NULL,
  COST_PRICE NUMBER(6) ,
  SELLING_PRICE NUMBER(6),
  );
/
CREATE TABLE product_brand(
  brand varchar(15),
  p_id varchar(10),
  s_no number(15) PRIMARY KEY,
  unit varchar2(25) not null,
  cost_price number(12),
  selling_price number(12)
 );
 /
 CREATE TABLE supplier_detail(
  supplier_id varchar(10) PRIMARY KEY,
  phone_no varchar(13) NOT NULL,
  email varchar(64)  NOT NULL,
  company_name varchar(20) NOT NULL,
  gstin_no varchar(15) NOT NULL,
  fax_no varchar(15) NOT NULL,
  address varchar(50) NOT NULL,
  supplier_name varchar2(30) not null
);
/
CREATE TABLE supplier_product(
  s_no varchar(10) PRIMARY KEY,
  s_id varchar(10) REFERENCES supplier_detail(supplier_id),
  p_id varchar(10),
  product_name varchar(15) ,
  gst number(5)
);
/
CREATE TABLE supplier_account(
  s_no varchar(10) PRIMARY KEY,
  supplier_id varchar(10) REFERENCES supplier_detail(supplier_id),
  account_no varchar(19) NOT NULL,
  account_holder_name varchar(30) NOT NULL,
  ifsc_code varchar(10) NOT NULL,
  bank_name varchar(20) NOT NULL
);
/
CREATE TABLE supplier_product_brand(
 brand varchar2(15),
 product_id varchar2(10),
 sup_id varchar2(10) REFERENCES supplier_detail(supplier_id),
 s_no  varchar2(20),
 unit  varchar2(25)
 );
/
CREATE TABLE supplieraccountsno(
  S_NO VARCHAR2(10) 
);
/
CREATE TABLE brandsno(
 s_no varchar(10) PRIMARY KEY
 );
/
CREATE TABLE supplierbrandsno(
 s_no varchar(10) PRIMARY KEY
 );
/
CREATE TABLE customer_detail(
 customer_id char(8) primary key,
 customer_name varchar(30) not null,
 gstin char(12) not null,
 address varchar(60) not null,
 district varchar(30) not null,
 state varchar(30) not null,
 phone_no number(10) ,
 email varchar(30) ,
 account_no varchar(18),
 holder_name varchar(30),
 ifcs_code varchar(16),
 bank_name varchar(24)
);
/
CREATE TABLE order_detail(
 order_number varchar(10) primary key,
 order_date date ,
 s_id varchar2(10) references supplier_detail(supplier_id),
 delivery_date date,
 inv_status varchar2(15),
 advance_amount number(8),
 dues_amount number(8),
 total_amount number(8)
);
/
CREATE TABLE ordered_product(
 s_no varchar2(10) primary key,
 product_name varchar2(15) not null,
 brand varchar2(15) not null,
 unit number(25) not null,
 quantity number(5) not null,
 igst number(3) not null,
 s_id varchar2(10),
 p_id varchar2(10),
 order_no varchar2(10),
 rate number(8),
 total_amount number(8)
);
/
CREATE TABLE ordered_product_amount(
 order_number varchar(10) primary key,
 total_amount number(7) not null,
 paid_amount  number(7) not null,
 balance_amount number(7) not null
);
/
CREATE TABLE customer_order_detail(
 order_number varchar(10) primary key,
 order_date date ,
 delivery_date date not null,
 p_id varchar2(5),
 mode_of_payment char(6),
 advance_payment number(6) not null,
 dues  number(6),
 total number(6) not null,
 customer_name varchar2(30) not null,
 address varchar2(25),
 status varchar2(10),
 c_id char(8)
);
/
CREATE TABLE customer_ordered_product(
s_no varchar2(10) primary key,
product_name varchar2(30) not null,
brand varchar2(15) not null,
unit varchar(25) not null,
unit_price number(5) not null,
quantity number(5) not null,
gst number(3) not null,
tot_amount number(8) not null,
p_id varchar2(5),
ord_no varchar2(10)
);
/
CREATE TABLE purchase_invoice(
 invoice_no varchar(10) primary key,
 invoice_date date not null,
 bill_no varchar(10) not null,
 order_no varchar(10) not null,
 supplier_id varchar(10) not null
);
/
CREATE TABLE invoice_product_detail(
 sno varchar(10) primary key,
 invoice_no varchar(10) ,
 p_id varchar(10) not null,
 product_nm varchar(15) not null,
 brand varchar(15) not null,
 unit varchar(25) not null,
 unit_price number(5) not null,
 quantity number(5) not null,
 igst number(3) not null,
 amount number(6) not null,
 deliverd_qty number(5) not null,
 balance_qty number(5) not null,
 pur_amount number(8)
 );
/
CREATE TABLE invoice_product_amount(
 invpoice_no varchar(10) primary key,
 oreder_no varchar(10) not null,
 total_amt number(6) not null,
 paid_amt number(6) not null,
 balance_amt number(6) not null,
 bill_no varchar(10) not null,
 supplier_id varchar(10) not null
);
/
CREATE TABLE invoice_detail(
 invoice_no char(8) primary key,
 invoice_date date not null,
 order_date date,
 mode_of_payment varchar2(6),
 advance_pay number(5),
 dues number(5),
 pay_amount number(6),
 customer_id char(10),
 order_no varchar2(12)
 );
 /
CREATE TABLE sell_product(
 sn_no number(4) primary key,
 description varchar(30) not null,
 unit_price number(6) not null,
 qty number(8) not null ,
 net_amount number(8) not null,
 tax_amount number(8) not null,
 tot_amount number(10) not null,
 invoice_no char(10),
 product_id varchar2(12),
 brand varchar2(20),
 unit varchar2(20)
);
/
CREATE TABLE stock_detail(
 stock_no varchar(10) primary key,
 invoice_no varchar(10),
 invoice_dt date,
 product_id varchar(10) not null,
 product_nm varchar(15) not null,
 unit varchar(25) not null,
 avl_quantity number(5),
 stock_lim number(5),
 brand varchar(15) not null
);
/
CREATE TABLE purchase_return(
 return_no varchar(10) primary key,
 invoice_date date not null, 
 bill_no varchar(10) not null,
 order_no varchar(10) not null,
 supplier_id varchar(10) not null,
 return_date date not null,
 invoice_no varchar2(10),
 tot_amount number(10)
);
/
CREATE TABLE purchase_return_product(
 s_no varchar(10) primary key,
 return_no varchar(10) references purchase_return(return_no),
 product_id varchar(10) not null,
 product_nm varchar(15) NOT NULL,
 brand varchar(10) not null,
 unit varchar(25) NOT NULL,
 unit_price number(5) NOT NULL,
 quantity number(5) not null,
 exp_date date not null,
 description varchar(20) not null,
 amount number(10)
);
/
CREATE TABLE sell_return(
 return_no varchar2(10) primary key,
 invoice_date date not null,
 bill_no varchar2(10) not null,
 order_no varchar2(10) not null,
 supplier_id varchar2(10),
 return_date date ,
 invoice_no varchar2(10),
 tot_amount number(10)
 );
 /
 CREATE TABLE sell_return_product(
 s_no varchar2(10) primary key,
 return_no varchar2(10),
 product_id varchar2(10) not null,
 product_nm varchar2(15) not null,
 brand varchar2(10) not null,
 unit varchar2(25) not null,
 unit_price number(5) not null,
 quantity number(5) not null,
 exp_date date not null,
 description varchar2(20) not null,
 amount number(10) 
);
/
CREATE TABLE contact(
 name varchar2(25),
 email varchar2(30),
 msg varchar2(100)
);
/
CREATE TABLE prpsno(
  s_no varchar2(10) not null 
);
/
CREATE TABLE ordersno(
  s_no  varchar2(10) not null
);