awk -v OFS=, '
    BEGIN {
        PROCINFO["FS"] = ","
    }

    NR == 1 && FNR == 1 {
        file = FILENAME
        sub(/.csv$/, "", file)
        dir = dirname(file)
        print "directory", "filename", $0
    }

    NR > 1 && FNR > 1 {
        file = FILENAME
        sub(/.csv$/, "", file)
        dir = dirname(file)
        print dir, file, $0
    }
' *.csv > merged.csv
