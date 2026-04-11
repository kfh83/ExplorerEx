//
// Debug Helpers
// 
// Copyright (c) 2015 - 2017 Sean Farrell <sean.farrell@rioki.org>
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
// 

#ifndef _DBG_H_
#define _DBG_H_

#include <windows.h>
#include <stdio.h>
#include <intrin.h>
#include <dbghelp.h>
#include <vector>
#include <string>
#include <sstream>
#include <iostream>

#pragma comment(lib, "dbghelp.lib")

#define DBG_TRACE(MSG, ...)  ::dbg::trace(MSG, __VA_ARGS__)

#define DBG_SOFT_ASSERT(COND) if ((COND) == false) { \
                                  DBG_TRACE(__FUNCTION__ ": Assertion '" #COND "' failed!\n"); \
                              }

#define DBG_ASSERT(COND) if ((COND) == false) { \
                            DBG_TRACE(__FUNCTION__ ": Assertion '" #COND "' failed!\n"); \
                            ::dbg::handle_assert(__FUNCTION__, #COND); \
                         }

#define DBG_FAIL(MSG) DBG_TRACE(__FUNCTION__ MSG "\n"); \
                      ::dbg::fail(__FUNCTION__, MSG);

namespace dbg
{
	inline
		void trace(const char* msg, ...)
	{
		char buff[1024];

		va_list args;
		va_start(args, msg);
		vsnprintf(buff, 1024, msg, args);

		OutputDebugStringA(buff);

		va_end(args);
	}

	inline
		std::string basename(const std::string& file)
	{
		size_t i = file.find_last_of("\\/");
		if (i == std::string::npos)
		{
			return file;
		}
		else
		{
			return file.substr(i + 1);
		}
	}

	struct StackFrame
	{
		DWORD64 address;
		std::string name;
		std::string module;
		unsigned int line;
		std::string file;
	};

	inline
		std::vector<StackFrame> stack_trace()
	{
#if _WIN64
		DWORD machine = IMAGE_FILE_MACHINE_AMD64;
#else
		DWORD machine = IMAGE_FILE_MACHINE_I386;
#endif
		HANDLE process = GetCurrentProcess();
		HANDLE thread = GetCurrentThread();

		if (SymInitialize(process, NULL, TRUE) == FALSE)
		{
			DBG_TRACE(__FUNCTION__ ": Failed to call SymInitialize.");
			return std::vector<StackFrame>();
		}

		SymSetOptions(SYMOPT_LOAD_LINES);

		CONTEXT    context = {};
		context.ContextFlags = CONTEXT_FULL;
		RtlCaptureContext(&context);

#if _WIN64
		STACKFRAME frame = {};
		frame.AddrPC.Offset = context.Rip;
		frame.AddrPC.Mode = AddrModeFlat;
		frame.AddrFrame.Offset = context.Rbp;
		frame.AddrFrame.Mode = AddrModeFlat;
		frame.AddrStack.Offset = context.Rsp;
		frame.AddrStack.Mode = AddrModeFlat;
#else
		STACKFRAME frame = {};
		frame.AddrPC.Offset = context.Eip;
		frame.AddrPC.Mode = AddrModeFlat;
		frame.AddrFrame.Offset = context.Ebp;
		frame.AddrFrame.Mode = AddrModeFlat;
		frame.AddrStack.Offset = context.Esp;
		frame.AddrStack.Mode = AddrModeFlat;
#endif


		bool first = true;

		std::vector<StackFrame> frames;
		while (StackWalk(machine, process, thread, &frame, &context, NULL, SymFunctionTableAccess, SymGetModuleBase, NULL))
		{
			StackFrame f = {};
			f.address = frame.AddrPC.Offset;

#if _WIN64
			DWORD64 moduleBase = 0;
#else
			DWORD moduleBase = 0;
#endif

			moduleBase = SymGetModuleBase(process, frame.AddrPC.Offset);

			char moduelBuff[MAX_PATH];
			if (moduleBase && GetModuleFileNameA((HINSTANCE)moduleBase, moduelBuff, MAX_PATH))
			{
				f.module = basename(moduelBuff);
			}
			else
			{
				f.module = "Unknown Module";
			}
#if _WIN64
			DWORD64 offset = 0;
#else
			DWORD offset = 0;
#endif
			char symbolBuffer[sizeof(IMAGEHLP_SYMBOL) + 255];
			PIMAGEHLP_SYMBOL symbol = (PIMAGEHLP_SYMBOL)symbolBuffer;
			symbol->SizeOfStruct = (sizeof IMAGEHLP_SYMBOL) + 255;
			symbol->MaxNameLength = 254;

			if (SymGetSymFromAddr(process, frame.AddrPC.Offset, &offset, symbol))
			{
				f.name = symbol->Name;
			}
			else
			{
				DWORD error = GetLastError();
				DBG_TRACE(__FUNCTION__ ": Failed to resolve address 0x%X: %u\n", frame.AddrPC.Offset, error);
				f.name = "Unknown Function";
			}

			IMAGEHLP_LINE line;
			line.SizeOfStruct = sizeof(IMAGEHLP_LINE);

			DWORD offset_ln = 0;
			if (SymGetLineFromAddr(process, frame.AddrPC.Offset, &offset_ln, &line))
			{
				f.file = line.FileName;
				f.line = line.LineNumber;
			}
			else
			{
				DWORD error = GetLastError();
				DBG_TRACE(__FUNCTION__ ": Failed to resolve line for 0x%X: %u\n", frame.AddrPC.Offset, error);
				f.line = 0;
			}

			if (!first)
			{
				frames.push_back(f);
			}
			first = false;
		}

		SymCleanup(process);

		return frames;
	}

	inline
		void handle_assert(const char* func, const char* cond)
	{
		std::stringstream buff;
		buff << func << ": Assertion '" << cond << "' failed! \n";
		buff << "\n";

		std::vector<StackFrame> stack = stack_trace();
		buff << "Callstack: \n";
		for (unsigned int i = 0; i < stack.size(); i++)
		{
			buff << "0x" << std::hex << stack[i].address << ": " << stack[i].name << "(" << std::dec << stack[i].line << ") in " << stack[i].module << "\n";
		}

		MessageBoxA(NULL, buff.str().c_str(), "Assert Failed", MB_OK | MB_ICONSTOP);
		abort();
	}

	inline
		void fail(const char* func, const char* msg)
	{
		std::stringstream buff;
		buff << func << ":  General Software Fault: '" << msg << "'! \n";
		buff << "\n";

		std::vector<StackFrame> stack = stack_trace();
		buff << "Callstack: \n";
		for (unsigned int i = 0; i < stack.size(); i++)
		{
			buff << "0x" << std::hex << stack[i].address << ": " << stack[i].name << "(" << stack[i].line << ") in " << stack[i].module << "\n";
		}

		MessageBoxA(NULL, buff.str().c_str(), "General Software Fault", MB_OK | MB_ICONSTOP);
		abort();
	}

	inline void printstacktrace()
	{
		std::stringstream buff;

		std::vector<StackFrame> stack = stack_trace();
		buff << "Callstack: \n";
		for (unsigned int i = 0; i < stack.size(); i++)
		{
			buff << "0x" << std::hex << stack[i].address << ": " << stack[i].name << "(" << stack[i].line << ") in " << stack[i].module << "\n";
		}

		std::cout << buff.str() << std::endl;
	}
}

#endif