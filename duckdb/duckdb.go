package duckdb

import (
	"database/sql"
	"errors"
	"fmt"
	"log"

	_ "github.com/marcboeker/go-duckdb/v2"
)

func ConnectDB() {
	// connector, err := db.NewConnector("test.db", nil)
	// if err != nil {
	// 	fmt.Println(err)
	// }
	// conn, err := connector.Connect(context.Background())
	// if err != nil {
	// 	fmt.Println(err)
	// }
	// defer conn.Close()

	db, err := sql.Open("duckdb", "./duckdb.db")
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	_, err = db.Exec(`CREATE TABLE people (id INTEGER, name VARCHAR)`)
	if err != nil {
		log.Fatal(err)
	}
	_, err = db.Exec(`INSERT INTO people VALUES (42, 'John')`)
	if err != nil {
		log.Fatal(err)
	}

	var (
		id   int
		name string
	)
	row := db.QueryRow(`SELECT id, name FROM people`)
	err = row.Scan(&id, &name)
	if errors.Is(err, sql.ErrNoRows) {
		log.Println("no rows")
	} else if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("id: %d, name: %s\n", id, name)
}
