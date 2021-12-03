package xlsx

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestEncrypt(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "encryptSHA1.xlsx"), Options{Password: "password"})
	assert.NoError(t, err)
	assert.EqualError(t, f.SaveAs(filepath.Join("test", "BadEncrypt.xlsx"), Options{Password: "password"}), "not support encryption currently")
	assert.NoError(t, f.Close())
}

func TestEncryptionMechanism(t *testing.T) {
	mechanism, err := encryptionMechanism([]byte{3, 0, 3, 0})
	assert.Equal(t, mechanism, "extensible")
	assert.EqualError(t, err, "unsupport encryption mechanism")
	_, err = encryptionMechanism([]byte{})
	assert.EqualError(t, err, "unknown encryption mechanism")
}

func TestHashing(t *testing.T) {
	assert.Equal(t, hashing("unsupportHashAlgorithm", []byte{}), []uint8([]byte(nil)))
}
