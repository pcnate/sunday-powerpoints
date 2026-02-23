-- Migration: 005_create_books_table
-- Description: Create books table for managing song book/source categories.
-- Books have an enabled flag so they can be hidden from dropdowns without deletion.

CREATE TABLE books (
  id         INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  name       VARCHAR(255) NOT NULL UNIQUE,
  enabled    BOOLEAN NOT NULL DEFAULT TRUE,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
