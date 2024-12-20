package main

import (
	"database/sql"
	"fmt"
	"log"

	_ "github.com/lib/pq"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Configurações do banco de dados
	connStr := "host=  port=  user=  password=  dbname=  sslmode=disable"
	db, err := sql.Open("postgres", connStr)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	// Executa a consulta SQL
	rows, err := db.Query(

		`SELECT  
    c.codcli, 
    c.nomcli, 
    c.nif, 
    COALESCE(p.despai, ' ') AS despai,
    cp.desconpag,
    v.nomven
FROM 
    public.clientes c
LEFT JOIN     
    public.paises p ON c.codpai = p.codpai
INNER JOIN 
    public.condicoespagamento cp ON c.codconpag = cp.codconpag
INNER JOIN 
    public.clienteenderecos ce ON c.codcli = ce.codcli -- Corrigido para usar c.codcli
INNER JOIN 
    public.vendedores v ON ce.codven = v.codven -- Corrigido para usar c.codven
WHERE     
c.codcli > 0
ORDER BY     
    c.codcli;`)
	if err != nil {
		log.Fatal(err)
	}
	defer rows.Close()

	// Cria um novo arquivo Excel
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.NewSheet(sheet)

	// Escreve o cabeçalho
	headers := []string{"codcli", "nomcli", "nif", "despai", "desmetpag", "nomven"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string('A'+i))
		f.SetCellValue(sheet, cell, header)
	}

	// Preenche o conteúdo
	rowIdx := 2
	for rows.Next() {
		var codcli int
		var nomcli, nif, despai, desmetpag, nomven string
		if err := rows.Scan(&codcli, &nomcli, &nif, &despai, &desmetpag, &nomven); err != nil {
			log.Fatal(err)
		}

		f.SetCellValue(sheet, fmt.Sprintf("A%d", rowIdx), codcli)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", rowIdx), nomcli)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", rowIdx), nif)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", rowIdx), despai)
		f.SetCellValue(sheet, fmt.Sprintf("E%d", rowIdx), desmetpag)
		f.SetCellValue(sheet, fmt.Sprintf("F%d", rowIdx), nomven)
		rowIdx++
	}

	// Salva o arquivo Excel
	if err := f.SaveAs("clientes.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Arquivo Excel gerado com sucesso!")
}
