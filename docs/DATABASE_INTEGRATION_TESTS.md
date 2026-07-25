# MySQL and PostgreSQL Integration Checks

The release smoke tests are framework-neutral and do not require a live MySQL/PostgreSQL server.
For release proof, configure optional environment variables and run the database integration check.

## MySQL

```env
MNB_EXCEL_TEST_MYSQL_DSN=mysql:host=127.0.0.1;port=3306;dbname=mnb_excel_test;charset=utf8mb4
MNB_EXCEL_TEST_MYSQL_USERNAME=root
MNB_EXCEL_TEST_MYSQL_PASSWORD=secret
```

## PostgreSQL

```env
MNB_EXCEL_TEST_PGSQL_DSN=pgsql:host=127.0.0.1;port=5432;dbname=mnb_excel_test
MNB_EXCEL_TEST_PGSQL_USERNAME=postgres
MNB_EXCEL_TEST_PGSQL_PASSWORD=secret
```

Run:

```bash
php tools/database-integration-check.php
```

