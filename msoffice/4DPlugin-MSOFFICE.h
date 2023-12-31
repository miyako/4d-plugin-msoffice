/* --------------------------------------------------------------------------------
 #
 #	4DPlugin-MSOFFICE.h
 #	source generated by 4D Plugin Wizard
 #	Project : MSOFFICE
 #	author : miyako
 #	2023/12/11
 #  
 # --------------------------------------------------------------------------------*/

#ifndef PLUGIN_MSOFFICE_H
#define PLUGIN_MSOFFICE_H

#include <stdio.h>
#include <stdlib.h>
#include <string>
#include <locale>
#include <fstream>
#include <cybozu/mmap.hpp>
#include <cybozu/file.hpp>
#include <cybozu/atoi.hpp>
#include <cybozu/option.hpp>
#include "cfb.hpp"
#include "decode.hpp"
#include "encode.hpp"
#include "make_dataspace.hpp"
#ifdef _MSC_VER
#include <cybozu/string.hpp>
#endif

#include "4DPluginAPI.h"

#include "4DPlugin-JSON.h"

#include "C_BLOB.h"
#include "C_TEXT.h"

#define USE_OPENXLSX 1
#if USE_OPENXLSX
#include <OpenXLSX.hpp>
#endif

#if VERSIONMAC
#import <CoreFoundation/CoreFoundation.h>
#endif

#if VERSIONWIN
#include <Shlobj.h>
#endif

#include <json/json.h>

#define USE_OPC 1
#if USE_OPC
#include <opc/opc.h>
#include <sstream>
#endif

#include <filesystem>
#include <fstream>

#define USE_XLNT 0
#if USE_XLNT
#include <xlnt/xlnt.hpp>
#endif

#pragma mark -

static void Write_to_spreadsheet_xlnt(PA_PluginParameters params);

#if USE_OPENXLSX
static void Write_to_spreadsheet(PA_PluginParameters params);
static void Read_from_spreadsheet(PA_PluginParameters params);
#endif

static void Verify_office_document(PA_PluginParameters params);
static void Decode_office_document(PA_PluginParameters params);
static void Encode_office_document(PA_PluginParameters params);

typedef enum {
    encode_aes_128 = 0,
    encode_aes_256 = 1
}encode_aes_t;


#pragma mark DEV

static bool parseReferencesObject(PA_ObjectRef references, Json::Value *value);
static void copy_macro_parts(CUTF8String& infilename_u8, CUTF8String& outfilename_u8);

#endif /* PLUGIN_MSOFFICE_H */
