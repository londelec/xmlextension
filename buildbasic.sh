#!/bin/sh
# LibreOffice XML import/export extension build script
# Created By AK
# Copyright Londelec UK Ltd
# Revision V1.1 02/01/2021
VERSION="buildbasicVersion=1.1"


ownname=$(basename $0)		# This is only script name
workpath=$(pwd)
sourcesdir="./src/*"
manifest="./manifest.xml"
descfile="./src/description.xml"
javametadir="META-INF"
zipname="temp.zip"
targetname="XMLExtension_v"
targetext=".oxt"
failure=0
dvers=""
includedir="\
./description \
./images \
"


# Usage help function
usage_func()
{
	echo "Build LibreOffice XML import/export extension"
	echo "Usage: ${ownname}"
	echo
	echo
}

# Logger function
# $1 = string to be logged
log_func()
{
	if ! test "x$1" = x; then
		echo "$1"
		#echo "$1">>${logfile}
	fi
}

# Copy manifest.xml to META-INF
copy_manifest_func()
{
	if ! $(mkdir "${tempdir}/${javametadir}"); then	# Directory creation failed
		log_func "Error: Failed to create ${tempdir}/${javametadir}"
		return 1	# Failure
	fi
	if ! test -d "${tempdir}/${javametadir}"; then	# Directory doesn't exist
		log_func "Error: Failed to create ${tempdir}/${javametadir}"
		return 1	# Failure
	fi

	if ! $(cp "${manifest}" "${tempdir}/${javametadir}"); then
		log_func "Error: Failed to copy ${manifest} to ${tempdir}/${javametadir}"
		return 1	# Failure
	fi
	return 0	# Success
}

# Extract version number from description.xml
get_extversion_func()
{
	# Ensure manifest exists
	if ! test -f "${descfile}"; then	# Description file doesn't exist
		log_func "Error: ${descfile} doesn't exist"
		return 1	# Failure
	fi
	dvers=$(cat "${descfile}" | grep "<version value="\"[0-9.]*\""" | grep -o "[0-9.]*")
	#echo "Debug: $dvers"	# For Debug
	if test "x${dvers}" = x; then
		log_func "Error: can't find version in ${descfile}"
		return 1	# Failure
	fi
	targetname="${targetname}${dvers}${targetext}"
	return 0	# Success
}

# Copy source directory contents to temporary directory
copy_sources_func()
{
	if ! $(cp -r ${sourcesdir} "${tempdir}"); then
		log_func "Error: Failed to copy ${sourcesdir} to ${tempdir}"
		return 1	# Failure
	fi
	return 0	# Success
}

# Copy include directories to temporary directory
copy_includes_func()
{
	for cdir in ${includedir}
	do
		if ! test -d "${cdir}"; then	# Directory doesn't exist
			log_func "Error: Include direcotory ${cdir} doesn't exist"
			return 1	# Failure
		fi 
		if ! $(cp -r "${cdir}" "${tempdir}"); then
			log_func "Error: Failed to copy ${sourcesdir} to ${tempdir}"
			return 1	# Failure
		fi
	done
	return 0	# Success
}


# Ensure manifest exists
if ! test -f "${manifest}"; then	# Manifest doesn't exist
	log_func "Error: ${manifest} doesn't exist"
	#usage_func
	exit 1		# Failure
fi

if ! get_extversion_func; then
	exit 1		# Failure
fi

tempdir=$(mktemp -d)
if ! test -d "${tempdir}"; then		# Directory doesn't exist
	log_func "Error: Failed to create ${manifest}"
	exit 1		# Failure
fi

if ! copy_manifest_func; then
	failure=1
fi
if ! copy_sources_func; then
	failure=1
fi
if ! copy_includes_func; then
	failure=1
fi


if test ${failure} -eq 0; then
	#echo "Debug: creating ZIP"	# For Debug
	cd ${tempdir}
	zip -r ${zipname} *
	#echo "Debug: ${targetname}"	# For Debug
	if ! $(mv "${tempdir}/${zipname}" "${workpath}/${targetname}"); then
		log_func "Error: mv ${tempdir}/${zipname} to ${workpath}/${targetname} failed"
	fi
fi


rm -rf "${tempdir}"
exit 0

