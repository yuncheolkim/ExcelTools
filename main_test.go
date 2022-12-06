package main

import (
	"encoding/json"
	"testing"
)

func TestJson(t *testing.T) {

	var cell interface{}
	e := json.Unmarshal([]byte(`[7000001,7000002,7000003,7000008,7000011,7000020,7000022,7000028,7000031,7000032]`), &cell)

	t.Log(e)
}
