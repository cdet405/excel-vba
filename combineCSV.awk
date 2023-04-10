awk -v OFS=, '
    NR == 1 && FNR == 1 {
        file = FILENAME
        sub(/.csv$/, "", file)
        print "filename", $0
    }
    NR > 1 && FNR > 1{
        file = FILENAME
        sub(/.csv$/, "", file)
        print file, $0
    }
