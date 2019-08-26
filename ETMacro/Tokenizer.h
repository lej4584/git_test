///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (c) 2011 Haangilsoft, Ltd.
// #203, 204, Jang YoungShill Hall, 1688-5 Shinil-dong, Daeduk-gu, Daejeon Metropolitan City, Korea, 306-230
// This code may not by copied or distributed or reused without permission from Haangilsoft.
//
//	Jinhong Kim
//	windgram@lollab.com
//	<REMARK>
//
//	TestTokenChar() returns: 
//		1: group maker, 2: separator, 3: delimiter, 4: null space, 0 for no-match.
//
//	If callback function receives of '9', it means that the token ends by separator character.
//
//
//
//	Revision History:
//	UPDATE DATE		WHO			ACT			DESCRIPTION
//	----------------------------------------------------------------------------------------------------------
//	2005-Sep-21		JHKIM		add			void InitDirectMemoryAccess( char* ptr, int len );
//	2005-Sep-21		JHKIM		add			int	TestTokenChar( char c );
//	2005-Dec-08		JHKIM		rename		Renamed InitDirectMemoryAccess() to SetLine()
//	2005-Dec-08		JHKIM		add			Added following APIs:
//											void ClearGroupMakerKey();
//											void ClearSeparatorKey();
//											void ClearDelimiterKey();
//											BOOL AddGroupMakerKey( char ch );
//											BOOL AddSeparatorKey( char ch );
//											BOOL AddDelimiterKey( char ch );
//											void SetKeyActivation( int type, BOOL b_state );
//											void LoadDefaultKey();
//											void ClearIgnoreKey();
//											void AddIgnoreKey( char ch );
//											BOOL AddNullKey( char ch );
//											void ClearNullKey();
//	2005-Dec-13		JHKIM		rename		PeekNextToken() -> GetNext()
//	2005-Dec-13		JHKIM		add			added: PeekNext()
//	2005-Dec-13		JHKIM		REMARK		Now, Tokenizer does not use linked-list equipment anymore.
//	2005-Dec-15		JHKIM		fix			fix to recognize group key as: "(double quote) and '(quote) have different meaning.
//								REMARK		A group opened by "(double quote) cannot be closed by '(quote) and vice versa.
//	2005-Dec-15		JHKIM		fix			fix the delimiter key processing error.
//	2005-Dec-15		JHKIM		add			added: SetCallback()
//	2005-Dec-15		JHKIM		change		API argument changed to: TestTokenChar( char c, BOOL b_cb )
//	2006-Jan-10		JHKIM		change		AddGroupMakerKey -> AddGroupMakerKey( char k, char k_sym )
//								REMARK		AddGroupMakerKey now supports the range definition with different opening/closing characters.
//	2006-Jan-10		JHKIM		add			GetGroupSymmetryKey( char k )
//	2006-Oct-20		JHKIM		fix			Error on generating token at EOF.
//	2007-May-07		JHKIM		add			PeekPtr() : returns the current position pointer.
//	2007-May-28		JHKIM		fix			GroupKey related: if the group key is exist in the delimiter list also, tokenizer must include it too.
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
#pragma once

#include "DataStream.h"

#define MAX_ITEM		50
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
namespace newSagaUtils
{
	class CTokenizer
	{
	public:
		CTokenizer(void);
		CTokenizer(char* filename);
		~CTokenizer(void);

	public:
		CDataStream ds;
		LONG fs;

		char buff[256];
		char group_maker_list[MAX_ITEM][2];
		int group_maker_list_cnt;
		char separator_list[MAX_ITEM];
		int separator_list_cnt;
		char delimiter_list[MAX_ITEM];
		int delimiter_list_cnt;
		char ignore_list[MAX_ITEM];
		int ignore_list_cnt;
		char null_list[MAX_ITEM];
		int null_list_cnt;

		BOOL b_groupmaker;
		BOOL b_separator;
		BOOL b_delimiter;

		void (*callbackfn)(int, char c, void*);
		void *cbparam;

	public:
		char* GetNext(int* len);
		char* PeekNext(int* len);
		char* PeekPtr(void);

		void Init(void);
		void SetLine(char* ptr, int len);
		int	TestTokenChar(char c, BOOL b_cb);

		void ClearGroupMakerKey(void);
		void ClearSeparatorKey(void);
		void ClearDelimiterKey(void);
		BOOL AddGroupMakerKey(char k, char k_sym);
		char GetGroupSymmetryKey(char k);
		BOOL AddSeparatorKey(char ch);
		BOOL AddDelimiterKey(char ch);
		void SetKeyActivation(int type, BOOL b_state);
		void LoadDefaultKey(void);
		void ClearIgnoreKey(void);
		BOOL AddIgnoreKey(char ch);
		BOOL AddNullKey(char ch);
		void ClearNullKey(void);
		void SetCallback(void (*_cb)(int, char c, void*), void* param);
	};
};
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////