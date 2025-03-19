#include "stdafx.h"
#include "script.h"
#include "globaldata.h"


Object *ScriptModule::sPrototype;


ResultType ScriptModule::Invoke(IObject_Invoke_PARAMS_DECL)
{
	auto var = aName ? mVars.Find(aName) : nullptr;
	if (!var && mIsBuiltinModule)
		// This is a slight hack to support built-in vars which haven't been referenced directly.
		var = g_script.FindOrAddBuiltInVar(aName, true, nullptr);
	if (!var || !var->IsExported())
		return ObjectBase::Invoke(IObject_Invoke_PARAMS);

	if (IS_INVOKE_SET && aParamCount == 1)
		return var->Assign(*aParam[0]);

	var->Get(aResultToken);
	if (aResultToken.Exited() || aParamCount == 0 && IS_INVOKE_GET)
		return aResultToken.Result();

	return Object::ApplyParams(aResultToken, aFlags, aParam, aParamCount);
}


ResultType Script::ParseModuleDirective(LPCTSTR aName)
{
	int at;
	if (mModules.Find(aName, &at))
		return ScriptError(ERR_DUPLICATE_DECLARATION, aName);
	// TODO: Validate module names.
	aName = SimpleHeap::Alloc(aName);
	auto mod = new ScriptModule(aName);
	// Let any previous #Warn settings carry over from the previous module, by default.
	mod->Warn_LocalSameAsGlobal = mCurrentModule->Warn_LocalSameAsGlobal;
	mod->Warn_Unreachable = mCurrentModule->Warn_Unreachable;
	mod->Warn_VarUnset = mCurrentModule->Warn_VarUnset;
	if (!mModules.Insert(mod, at))
		return MemoryError();
	CloseCurrentModule();
	mCurrentModule = mod;
	return CONDITION_TRUE;
}


ResultType Script::ParseImportStatement(LPTSTR aBuf)
{
	aBuf = omit_leading_whitespace(aBuf);
	auto imp = new ScriptImport();
	imp->names = SimpleHeap::Alloc(aBuf);
	imp->mod = nullptr;
	imp->line_number = mCombinedLineNumber;
	imp->file_index = mCurrFileIndex;
	imp->next = mCurrentModule->mImports;
	mCurrentModule->mImports = imp;
	return OK;
}


Var *Script::FindImportedVar(LPCTSTR aVarName)
{
	for (auto imp = CurrentModule()->mImports; imp; imp = imp->next)
	{
		if (*imp->names == '*' && imp->mod) // mod can be null during DerefInclude().
		{
			auto var = imp->mod->mVars.Find(aVarName);
			if (var && var->IsExported())
				return var;
		}
	}
	return nullptr;
}


// Add a new Var to the current module.
// Raises an error if a conflicting declaration exists.
// May use an existing Var if not previously marked as declared, such as if created by Export.
// Caller provides persistent memory for aVarName.
Var *Script::AddNewImportVar(LPTSTR aVarName)
{
	int at;
	auto var = mCurrentModule->mVars.Find(aVarName, &at);
	if (var)
	{
		// Only declared, assigned or exported variables should exist at this point.
		if (var->IsDeclared() || var->IsAssignedSomewhere())
		{
			ConflictingDeclarationError(_T("Import"), var);
			return nullptr;
		}
		var->Scope() |= VAR_DECLARED;
	}
	else
		var = new Var(aVarName, VAR_DECLARE_GLOBAL);
	if (!mCurrentModule->mVars.Insert(var, at))
	{
		delete var;
		MemoryError();
		return nullptr;
	}
	return var;
}


ResultType Script::ResolveImports()
{
	for (mCurrentModule = mLastModule; mCurrentModule; mCurrentModule = mCurrentModule->mPrev)
	{
		for (auto imp = mCurrentModule->mImports; imp; imp = imp->next)
		{
			if (!imp->mod && !ResolveImports(*imp))
				return FAIL;
		}
	}
	return OK;
}


ResultType Script::ResolveImports(ScriptImport &imp)
{
	// Set early for code size, in case of error.
	mCurrLine = nullptr;
	mCombinedLineNumber = imp.line_number;
	mCurrFileIndex = imp.file_index;

	bool import_file = false;
	LPTSTR cp = imp.names, mod_name, var_name = nullptr;
	if (*cp == '{' || *cp == '*')
	{
		if (*cp == '{')
			cp = _tcschr(cp, '}'); // Should always be found due to GetLineContExpr().
		cp = omit_leading_whitespace(cp + 1);
		if (_tcsnicmp(cp, _T("From"), 4) || !IS_SPACE_OR_TAB(cp[4]))
			return ScriptError(_T("Invalid import"), imp.names);
		cp = omit_leading_whitespace(cp + 5);
	}
	if (import_file = (*cp == '"' || *cp == '\''))
	{
		mod_name = cp + 1;
		cp += FindTextDelim(cp, *cp, 1);
		if (!*cp || cp[1] && !IS_SPACE_OR_TAB(cp[1]))
			return ScriptError(_T("Invalid import"), imp.names);
		*cp++ = '\0';
		ConvertEscapeSequences(mod_name);
  	}
	else
	{
		mod_name = cp;
		while (*cp && !IS_SPACE_OR_TAB(*cp)) ++cp;
		if (*cp)
			*cp++ = '\0';
	}
	if (*cp) // There's a character after the quote or space which terminates the module name/path.
	{
		cp = omit_leading_whitespace(cp);
		if (_tcsnicmp(cp, _T("as"), 2) || !IS_SPACE_OR_TAB(cp[2]))
			return ScriptError(_T("Invalid import"), cp);
		var_name = omit_leading_whitespace(cp + 3);
	}
	else if (mod_name == imp.names) // `Import M`, not `Import {} from M` or `Import "file"`.
		var_name = mod_name;

	int at;
	if (  !(imp.mod = mModules.Find(mod_name, &at))  )
	{
		FileIndexType file_index;
		switch (FindModuleFileIndex(mod_name, file_index))
		{
		default:	return ScriptError(_T("Module not found"), mod_name);
		case FAIL:	return FAIL;
		case OK:	break;
		}
		// Search by file index in case of two different paths referring to the same file.
		for (int i = 0; i < mModules.mCount; ++i)
			if (mModules.mItem[i]->mSelfFileIndex == file_index)
			{
				imp.mod = mModules.mItem[i];
				break;
			}
		if (!imp.mod)
		{
			auto path = Line::sSourceFile[file_index];
			auto cur_mod = mCurrentModule;
			auto last_mod = mLastModule;
			mLastModule = nullptr; // Start a new chain.
			imp.mod = mCurrentModule = new ScriptModule(mod_name);
			imp.mod->mSelfFileIndex = file_index;
			if (!mModules.Insert(imp.mod, at))
				return MemoryError();
			if (!LoadIncludedFile(path, false, false))
				return FAIL;
			if (!CloseCurrentModule())
				return FAIL;
			if (!ResolveImports()) // Resolve imports in all modules that were just included.
				return FAIL;
			imp.mod->mPrev = last_mod; // Join to previous chain.
			mCurrentModule = cur_mod;
		}
	}

	if (var_name)
	{
		// Do not reuse mSelf or a previous Var created by an import even if mod_name == var_name,
		// since the exported status of the Var (VAR_EXPORTED) shouldn't propagate between modules.
		auto var = AddNewImportVar(var_name);
		if (!var)
			return FAIL;
		if (imp.mod->mSelf) // Default export.
		{
			// For code size, aliasing is used even for constants.  For non-dynamic references to
			// constants, PreparseVarRefs() eliminates both the alias and the var reference itself.
			var->UpdateAlias(imp.mod->mSelf);
		}
		else
		{
			var->Assign(imp.mod);
			var->MakeReadOnly();
		}
	}

	if (*imp.names == '{')
	{
		for (cp = imp.names + 1; *(cp = omit_leading_whitespace(cp)) != '}'; ++cp)
		{
			auto c = *(cp = find_identifier_end(var_name = mod_name = cp));
			*cp = '\0';
			while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
			auto exported = imp.mod->mVars.Find(mod_name);
			if (!exported)
				return ScriptError(_T("No such export"), mod_name);
			if (!_tcsnicmp(cp, _T("as"), 2) && IS_SPACE_OR_TAB(cp[2]))
			{
				var_name = omit_leading_whitespace(cp + 3);
				cp = find_identifier_end(var_name);
				c = *cp;
				*cp = '\0';
				while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
			}
			if (!(c == ',' || c == '}'))
			{
				*cp = c;
				return ScriptError(_T("Invalid import"), cp);
			}
			auto imported = AddNewImportVar(var_name);
			if (!imported)
				return FAIL;
			imported->UpdateAlias(exported); // See the other UpdateAlias call above for comments.
			if (c == '}')
				break;
		}
	}

	return OK;
}


ResultType Script::CloseCurrentModule()
{
	// Terminate each module with RETURN so that all labels have a target and
	// all control flow statements that need it have a non-null mRelatedLine.
	if (!AddLine(ACT_EXIT))
		return FAIL;

	mCurrentModule->mPrev = mLastModule;
	mLastModule = mCurrentModule;

	ASSERT(!mCurrentModule->mFirstLine || mCurrentModule->mFirstLine == mFirstLine);
	mCurrentModule->mFirstLine = mFirstLine;
	mFirstLine = nullptr; // Start a new linked list of lines.
	mLastLine = nullptr;
	mLastLabel = nullptr; // Start a new linked list of labels.
	return OK;
}


bool ScriptModule::HasFileIndex(FileIndexType aFile)
{
	for (int i = 0; i < mFilesCount; ++i)
		if (mFiles[i] == aFile)
			return true;
	return false;
}


ResultType ScriptModule::AddFileIndex(FileIndexType aFile)
{
	if (mFilesCount == mFilesCountMax)
	{
		auto new_size = mFilesCount ? mFilesCount * 2 : 8;
		auto new_files = (FileIndexType*)realloc(mFiles, new_size * sizeof(FileIndexType));
		if (!new_files)
			return MemoryError();
		mFiles = new_files;
		mFilesCountMax = new_size;
	}
	mFiles[mFilesCount++] = aFile;
	return OK;
}


IObject *ScriptModule::FindGlobalObject(LPCTSTR aName)
{
	if (auto var = mVars.Find(aName))
	{
		ASSERT(!var->IsVirtual());
		return var->ToObject();
	}
	return nullptr;
}


LPCWSTR Script::InitModuleSearchPath()
{
	SetCurrentDirectory(mFileDir);

	// The size of this array sets the total limit for the entire list of directories, resolved to full paths.
	// This is enough for extreme cases, up to the total limit for all environment variables (32767 chars).
	TCHAR buf[MAX_WIDE_PATH], var_buf[MAX_WIDE_PATH];
	LPTSTR d, cp = buf;
	LPCTSTR path_spec;
	DWORD len, space;

	if (GetEnvironmentVariable(L"AhkImportPath", var_buf, _countof(var_buf)))
		path_spec = var_buf;
	else
		path_spec = _T(".;%A_MyDocuments%\\AutoHotkey;%A_AhkPath%\\..");

	LPTSTR deref_path;
	if (!DerefInclude(deref_path, path_spec))
		return _T("");

	for (auto p = deref_path; *p; p = d + 1)
	{
		while (*p == ';') ++p;
		for (d = p; *d && *d != ';'; ++d);
		if (d == p // Empty item at the end.
			|| (cp - buf) + (d - p) + 1 > _countof(buf)) // Due to rarity, ignore any items that won't fit.
			break;
		*d = '\0'; // Terminate within deref_path.
		space = (DWORD)(_countof(buf) - (cp - buf));
		len = GetFullPathName(p, space, cp, nullptr);
		if (!len || len >= space)
			continue; // Ignore this item.
		cp += len + 1; // Write the next item after the null terminator.
	}
	if (cp == buf) // No items; could happen in the case of a malformed env var.
		*cp++ = '\0'; // Ensure double-null termination.
	
	free(deref_path);
	return SimpleHeap::Alloc(buf, cp - buf);
}


ResultType Script::FindModuleFileIndex(LPCTSTR aName, FileIndexType &aFileIndex)
{
	static auto search_path = InitModuleSearchPath();
	const auto suffix = EXT_AUTOHOTKEY;
	const auto suffix_length = _countof(EXT_AUTOHOTKEY) - 1;
	const auto dir_suffix = _T("\\__Init") EXT_AUTOHOTKEY;
	const auto dir_suffix_length = _countof(_T("\\__Init") EXT_AUTOHOTKEY) - 1;

	TCHAR buf[T_MAX_PATH];
	
	auto name_length = _tcslen(aName);
	if (name_length + dir_suffix_length >= _countof(buf))
		return CONDITION_FALSE;

	DWORD attr;
	size_t dir_length;
	for (auto dir = search_path; *dir; dir += dir_length + 1)
	{
		dir_length = _tcslen(dir);
		if (!SetCurrentDirectory(dir))
			continue; // Ignore this item.

		tmemcpy(buf, aName, name_length + 1);

		// Try exact aName first.
		attr = GetFileAttributes(buf);
		auto p = buf + name_length;
		if ((attr & FILE_ATTRIBUTE_DIRECTORY) && attr != INVALID_FILE_ATTRIBUTES)
		{
			// Try aName\__module.ahk.
			tmemcpy(p, dir_suffix, dir_suffix_length + 1);
			attr = GetFileAttributes(buf);
		}
		if (attr & FILE_ATTRIBUTE_DIRECTORY) // No file found yet.
		{
			// Try aName.ahk.
			tmemcpy(p, suffix, suffix_length + 1);
			attr = GetFileAttributes(buf);
		}
		if ((attr & FILE_ATTRIBUTE_DIRECTORY) == 0) // File exists and is not a directory.
			return SourceFileIndex(buf, aFileIndex);
	}
	return CONDITION_FALSE;
}
