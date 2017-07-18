#!/bin/sh

if [ $# != 2 ]
then
	echo "Usage:"
	echo " ./genStat.sh [YYYYMMDD] [YYYYMMDD]"
	echo ""
	exit 1
fi

TIMETABLE=(
		"00:00~00:30" "00:30~01:00" "01:00~01:30" "01:30~02:00"
		"02:00~02:30" "02:30~03:00" "03:00~03:30" "03:30~04:00"
		"04:00~04:30" "04:30~05:00" "05:00~05:30" "05:30~06:00"
		"06:00~06:30" "06:30~07:00" "07:00~07:30" "07:30~08:00"
		"08:00~08:30" "08:30~09:00" "09:00~09:30" "09:30~10:00"
		"10:00~10:30" "10:30~11:00" "11:00~11:30" "11:30~12:00"
		"12:00~12:30" "12:30~13:00" "13:00~13:30" "13:30~14:00"
		"14:00~14:30" "14:30~15:00" "15:00~15:30" "15:30~16:00"
		"16:00~16:30" "16:30~17:00" "17:00~17:30" "17:30~18:00"
		"18:00~18:30" "18:30~19:00" "19:00~19:30" "19:30~20:00"
		"20:00~20:30" "20:30~21:00" "21:00~21:30" "21:30~22:00"
		"22:00~22:30" "22:30~23:00" "23:00~23:30" "23:30~24:00"
		)

START_DATE=$1
END_DATE=$2
CUR_DATE=$START_DATE
FILENAME="stat_${START_DATE}-${END_DATE}.dat"


print_cycle() {
	TO_DATE=$1

	BASEDATE=$(LANG=C date --date "${TO_DATE}" "+%Y-%m-%d(%a)")

	printf "%s|" $BASEDATE

	for (( cycle=1; cycle <= 48; cycle++ ))
	do
		icycle=$((${cycle#0} - 1))
		cytime=`echo ${TIMETABLE[$icycle]}`

		if [ $cycle -ge 10 ]
		then
			BE="${TO_DATE}0${cycle}"
		else
			BE="${TO_DATE}00${cycle}"
		fi

		SAMR=`echo -e "hi -b ${BE} -e ${BE}\nq\n"|$NOCT_HOME/bin/noctconsole |grep "Completed Lines" |awk -F " " '{print $4}'`

		printf "%d," $cycle
		printf "%s," $cytime
		printf "%d|" $SAMR
	done

	printf "\n"
}

rm -f ./$FILENAME

while [[ $CUR_DATE -le $END_DATE ]]
do

	print_cycle $CUR_DATE >>./$FILENAME

	CUR_DATE=`date -d "${CUR_DATE} + 1 day" "+%Y%m%d"`
done
