/* Version: 16.0.8305.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

declare namespace OneNote {
    /**
     *
     * Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Application extends OfficeExtension.ClientObject {
        private m_notebooks;
        private m__platform;
        readonly _className: string;
        /**
         *
         * Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebooks: OneNote.NotebookCollection;
        readonly _platform: string;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebook(): OneNote.Notebook;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebookOrNull(): OneNote.Notebook;
        /**
         *
         * Gets the active outline if one exists, If no outline is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutline(): OneNote.Outline;
        /**
         *
         * Gets the active outline if one exists, otherwise returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutlineOrNull(): OneNote.Outline;
        /**
         *
         * Gets the active page if one exists. If no page is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePage(): OneNote.Page;
        /**
         *
         * Gets the active page if one exists. If no page is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePageOrNull(): OneNote.Page;
        /**
         *
         * Gets the active Paragraph if one exists, If no Paragraph is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi]
         */
        getActiveParagraph(): OneNote.Paragraph;
        /**
         *
         * Gets the active Paragraph if one exists, otherwise returns null.
         *
         * [Api set: OneNoteApi]
         */
        getActiveParagraphOrNull(): OneNote.Paragraph;
        /**
         *
         * Gets the active section if one exists. If no section is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSection(): OneNote.Section;
        /**
         *
         * Gets the active section if one exists. If no section is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSectionOrNull(): OneNote.Section;
        insertHtmlAtCurrentPosition(html: string): void;
        /**
         *
         * Opens the specified page in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param page The page to open.
         */
        navigateToPage(page: OneNote.Page): void;
        /**
         *
         * Gets the specified page, and opens it in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param url The client url of the page to open.
         */
        navigateToPageWithClientUrl(url: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Application;
        toJSON(): {};
    }
    /**
     *
     * Represents ink analysis data for a given set of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysis extends OfficeExtension.ClientObject {
        private m_id;
        private m_page;
        private m_paragraphs;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the parent page object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly page: OneNote.Page;
        /**
         *
         * Gets the ink analysis paragraphs in this page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.InkAnalysisParagraphCollection;
        /**
         *
         * Gets the ID of the InkAnalysis object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysis;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysis;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysis;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * Represents ink analysis data for an identified paragraph formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisParagraph extends OfficeExtension.ClientObject {
        private m_id;
        private m_inkAnalysis;
        private m_lines;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Reference to the parent InkAnalysisPage. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkAnalysis: OneNote.InkAnalysis;
        /**
         *
         * Gets the ink analysis lines in this ink analysis paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly lines: OneNote.InkAnalysisLineCollection;
        /**
         *
         * Gets the ID of the InkAnalysisParagraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisParagraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisParagraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisParagraph;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * Represents a collection of InkAnalysisParagraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisParagraphCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.InkAnalysisParagraph>;
        /**
         *
         * Returns the number of InkAnalysisParagraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a InkAnalysisParagraph object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the InkAnalysisParagraph object, or the index location of the InkAnalysisParagraph object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets a InkAnalysisParagraph on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisParagraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisParagraphCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents ink analysis data for an identified text line formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisLine extends OfficeExtension.ClientObject {
        private m_id;
        private m_paragraph;
        private m_words;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Reference to the parent InkAnalysisParagraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets the ink analysis words in this ink analysis line. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly words: OneNote.InkAnalysisWordCollection;
        /**
         *
         * Gets the ID of the InkAnalysisLine object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisLine;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisLine;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisLine;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * Represents a collection of InkAnalysisLine objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisLineCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.InkAnalysisLine>;
        /**
         *
         * Returns the number of InkAnalysisLines in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a InkAnalysisLine object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the InkAnalysisLine object, or the index location of the InkAnalysisLine object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisLine;
        /**
         *
         * Gets a InkAnalysisLine on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisLine;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisLineCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisLineCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisLineCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents ink analysis data for an identified word formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisWord extends OfficeExtension.ClientObject {
        private m_id;
        private m_languageId;
        private m_line;
        private m_strokePointers;
        private m_wordAlternates;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Reference to the parent InkAnalysisLine. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly line: OneNote.InkAnalysisLine;
        /**
         *
         * Gets the ID of the InkAnalysisWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The id of the recognized language in this inkAnalysisWord. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly strokePointers: Array<OneNote.InkStrokePointer>;
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly wordAlternates: Array<string>;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisWord;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisWord;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisWord;
        toJSON(): {
            "id": string;
            "languageId": string;
            "strokePointers": InkStrokePointer[];
            "wordAlternates": string[];
        };
    }
    /**
     *
     * Represents a collection of InkAnalysisWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkAnalysisWordCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.InkAnalysisWord>;
        /**
         *
         * Returns the number of InkAnalysisWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the InkAnalysisWord object, or the index location of the InkAnalysisWord object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisWord;
        /**
         *
         * Gets a InkAnalysisWord on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisWordCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisWordCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisWordCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a group of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class FloatingInk extends OfficeExtension.ClientObject {
        private m_id;
        private m_inkStrokes;
        private m_pageContent;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the strokes of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkStrokes: OneNote.InkStrokeCollection;
        /**
         *
         * Gets the PageContent parent of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the ID of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.FloatingInk;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.FloatingInk;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.FloatingInk;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * Represents a single stroke of ink.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkStroke extends OfficeExtension.ClientObject {
        private m_floatingInk;
        private m_id;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly floatingInk: OneNote.FloatingInk;
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkStroke;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkStroke;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkStroke;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * Represents a collection of InkStroke objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkStrokeCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.InkStroke>;
        /**
         *
         * Returns the number of InkStrokes in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a InkStroke object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the InkStroke object, or the index location of the InkStroke object in the collection.
         */
        getItem(index: number | string): OneNote.InkStroke;
        /**
         *
         * Gets a InkStroke on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkStroke;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkStrokeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkStrokeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkStrokeCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * A container for the ink in a word in a paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkWord extends OfficeExtension.ClientObject {
        private m_id;
        private m_languageId;
        private m_paragraph;
        private m_wordAlternates;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * The parent paragraph containing the ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the InkWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The id of the recognized language in this ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi]
         */
        readonly wordAlternates: Array<string>;
        readonly _ReferenceId: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkWord;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkWord;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkWord;
        toJSON(): {
            "id": string;
            "languageId": string;
            "wordAlternates": string[];
        };
    }
    /**
     *
     * Represents a collection of InkWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class InkWordCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.InkWord>;
        /**
         *
         * Returns the number of InkWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a InkWord object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the InkWord object, or the index location of the InkWord object in the collection.
         */
        getItem(index: number | string): OneNote.InkWord;
        /**
         *
         * Gets a InkWord on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkWordCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkWordCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkWordCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a OneNote notebook. Notebooks contain section groups and sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Notebook extends OfficeExtension.ClientObject {
        private m_baseUrl;
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_sectionGroups;
        private m_sections;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * The section groups in the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The the sections of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sections: OneNote.SectionCollection;
        /**
         *
         * The url of the site that this notebook is located. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly baseUrl: string;
        /**
         *
         * The client url of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the name of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        readonly _ReferenceId: string;
        /**
         *
         * Adds a new section to the end of the notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the new section.
         */
        addSection(name: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of the notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the new section.
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Notebook;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Notebook;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Notebook;
        toJSON(): {
            "baseUrl": string;
            "clientUrl": string;
            "id": string;
            "name": string;
        };
    }
    /**
     *
     * Represents a collection of notebooks.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class NotebookCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.Notebook>;
        /**
         *
         * Returns the number of notebooks in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets the collection of notebooks with the specified name that are open in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the notebook.
         */
        getByName(name: string): OneNote.NotebookCollection;
        /**
         *
         * Gets a notebook by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the notebook, or the index location of the notebook in the collection.
         */
        getItem(index: number | string): OneNote.Notebook;
        /**
         *
         * Gets a notebook on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Notebook;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.NotebookCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.NotebookCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.NotebookCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a OneNote section group. Section groups can contain sections and other section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class SectionGroup extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_notebook;
        private m_parentSectionGroup;
        private m_parentSectionGroupOrNull;
        private m_sectionGroups;
        private m_sections;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the notebook that contains the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebook: OneNote.Notebook;
        /**
         *
         * Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The collection of section groups in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The collection of sections in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sections: OneNote.SectionCollection;
        /**
         *
         * The client url of the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the name of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        readonly _ReferenceId: string;
        /**
         *
         * Adds a new section to the end of the section group.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title The name of the new section.
         */
        addSection(title: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of this sectionGroup.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the new section.
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionGroup;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionGroup;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionGroup;
        toJSON(): {
            "clientUrl": string;
            "id": string;
            "name": string;
        };
    }
    /**
     *
     * Represents a collection of section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class SectionGroupCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.SectionGroup>;
        /**
         *
         * Returns the number of section groups in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets the collection of section groups with the specified name.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the section group.
         */
        getByName(name: string): OneNote.SectionGroupCollection;
        /**
         *
         * Gets a section group by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the section group, or the index location of the section group in the collection.
         */
        getItem(index: number | string): OneNote.SectionGroup;
        /**
         *
         * Gets a section group on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionGroupCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionGroupCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionGroupCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a OneNote section. Sections can contain pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Section extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_notebook;
        private m_pages;
        private m_parentSectionGroup;
        private m_parentSectionGroupOrNull;
        private m_webUrl;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the notebook that contains the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebook: OneNote.Notebook;
        /**
         *
         * The collection of pages in the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pages: OneNote.PageCollection;
        /**
         *
         * Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section. Returns null if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The client url of the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the name of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        /**
         *
         * The web url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly webUrl: string;
        readonly _ReferenceId: string;
        /**
         *
         * Adds a new page to the end of the section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title The title of the new page.
         */
        addPage(title: string): OneNote.Page;
        /**
         *
         * Copies this section to specified notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationNotebook The notebook to copy this section to.
         */
        copyToNotebook(destinationNotebook: OneNote.Notebook): OneNote.Section;
        /**
         *
         * Copies this section to specified section group.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationSectionGroup The section group to copy this section to.
         */
        copyToSectionGroup(destinationSectionGroup: OneNote.SectionGroup): OneNote.Section;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a new section before or after the current section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param location The location of the new section relative to the current section.
         * @param title The name of the new section.
         */
        insertSectionAsSibling(location: string, title: string): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Section;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Section;
        toJSON(): {
            "clientUrl": string;
            "id": string;
            "name": string;
            "webUrl": string;
        };
    }
    /**
     *
     * Represents a collection of sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class SectionCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.Section>;
        /**
         *
         * Returns the number of sections in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets the collection of sections with the specified name.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name The name of the section.
         */
        getByName(name: string): OneNote.SectionCollection;
        /**
         *
         * Gets a section by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the section, or the index location of the section in the collection.
         */
        getItem(index: number | string): OneNote.Section;
        /**
         *
         * Gets a section on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Page extends OfficeExtension.ClientObject {
        private m_classNotebookPageSource;
        private m_clientUrl;
        private m_contents;
        private m_id;
        private m_inkAnalysisOrNull;
        private m_pageLevel;
        private m_parentSection;
        private m_title;
        private m_webUrl;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * The collection of PageContent objects on the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly contents: OneNote.PageContentCollection;
        /**
         *
         * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkAnalysisOrNull: OneNote.InkAnalysis;
        /**
         *
         * Gets the section that contains the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSection: OneNote.Section;
        /**
         *
         * Gets the ClassNotebookPageSource to the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly classNotebookPageSource: string;
        /**
         *
         * The client url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets or sets the indentation level of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageLevel: number;
        /**
         *
         * Gets or sets the title of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        title: string;
        /**
         *
         * The web url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly webUrl: string;
        readonly _ReferenceId: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.PageUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Page): void;
        /**
         *
         * Adds an Outline to the page at the specified position.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param left The left position of the top, left corner of the Outline.
         * @param top The top position of the top, left corner of the Outline.
         * @param html An HTML string that describes the visual presentation of the Outline. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         */
        addOutline(left: number, top: number, html: string): OneNote.Outline;
        /**
         *
         * Copies this page to specified section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationSection The section to copy this page to.
         */
        copyToSection(destinationSection: OneNote.Section): OneNote.Page;
        /**
         *
         * Copies this page to specified section and sets ClassNotebookPageSource.
         *
         * [Api set: OneNoteApi 1.1]
         */
        copyToSectionAndSetClassNotebookPageSource(destinationSection: OneNote.Section): OneNote.Page;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Does the page has content title.
         *
         * [Api set: OneNoteApi]
         */
        hasTitleContent(): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Inserts a new page before or after the current page.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param location The location of the new page relative to the current page.
         * @param title The title of the new page.
         */
        insertPageAsSibling(location: string, title: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Page;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Page;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Page;
        toJSON(): {
            "classNotebookPageSource": string;
            "clientUrl": string;
            "id": string;
            "pageLevel": number;
            "title": string;
            "webUrl": string;
        };
    }
    /**
     *
     * Represents a collection of pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class PageCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.Page>;
        /**
         *
         * Returns the number of pages in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets the collection of pages with the specified title.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title The title of the page.
         */
        getByTitle(title: string): OneNote.PageCollection;
        /**
         *
         * Gets a page by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the page, or the index location of the page in the collection.
         */
        getItem(index: number | string): OneNote.Page;
        /**
         *
         * Gets a page on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class PageContent extends OfficeExtension.ClientObject {
        private m_id;
        private m_image;
        private m_ink;
        private m_left;
        private m_outline;
        private m_parentPage;
        private m_top;
        private m_type;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly image: OneNote.Image;
        /**
         *
         * Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly ink: OneNote.FloatingInk;
        /**
         *
         * Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly outline: OneNote.Outline;
        /**
         *
         * Gets the page that contains the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentPage: OneNote.Page;
        /**
         *
         * Gets the ID of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets or sets the left (X-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        left: number;
        /**
         *
         * Gets or sets the top (Y-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        top: number;
        /**
         *
         * Gets the type of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly type: string;
        readonly _ReferenceId: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.PageContentUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: PageContent): void;
        /**
         *
         * Deletes the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageContent;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageContent;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageContent;
        toJSON(): {
            "id": string;
            "left": number;
            "top": number;
            "type": string;
        };
    }
    /**
     *
     * Represents the contents of a page, as a collection of PageContent objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class PageContentCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.PageContent>;
        /**
         *
         * Returns the number of page contents in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a PageContent object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the PageContent object, or the index location of the PageContent object in the collection.
         */
        getItem(index: number | string): OneNote.PageContent;
        /**
         *
         * Gets a page content on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.PageContent;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageContentCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageContentCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageContentCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a container for Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Outline extends OfficeExtension.ClientObject {
        private m_id;
        private m_pageContent;
        private m_paragraphs;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the collection of Paragraph objects in the Outline. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the ID of the Outline object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        readonly _ReferenceId: string;
        /**
         *
         * Adds the specified HTML to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param html The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param base64EncodedImage HTML string to append.
         * @param width Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param paragraphText HTML string to append.
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to the bottom of the outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        appendTable(rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         *
         * Check if the outline is title outline.
         *
         * [Api set: OneNoteApi 1.1]
         */
        isTitle(): OfficeExtension.ClientResult<boolean>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Outline;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Outline;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Outline;
        toJSON(): {
            "id": string;
        };
    }
    /**
     *
     * A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Paragraph extends OfficeExtension.ClientObject {
        private m_id;
        private m_image;
        private m_inkWords;
        private m_outline;
        private m_paragraphs;
        private m_parentParagraph;
        private m_parentParagraphOrNull;
        private m_parentTableCell;
        private m_parentTableCellOrNull;
        private m_richText;
        private m_table;
        private m_type;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly image: OneNote.Image;
        /**
         *
         * Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkWords: OneNote.InkWordCollection;
        /**
         *
         * Gets the Outline object that contains the Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly outline: OneNote.Outline;
        /**
         *
         * The collection of paragraphs under this paragraph. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent paragraph object. Throws if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentParagraph: OneNote.Paragraph;
        /**
         *
         * Gets the parent paragraph object. Returns null if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentParagraphOrNull: OneNote.Paragraph;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTableCell: OneNote.TableCell;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTableCellOrNull: OneNote.TableCell;
        /**
         *
         * Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly richText: OneNote.RichText;
        /**
         *
         * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly table: OneNote.Table;
        /**
         *
         * Gets the ID of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the type of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly type: string;
        readonly _ReferenceId: string;
        /**
         *
         * Deletes the paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         *
         * Get list information of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        getParagraphInfo(): OfficeExtension.ClientResult<OneNote.ParagraphInfo>;
        /**
         *
         * Inserts the specified HTML content
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation The location of new contents relative to the current Paragraph.
         * @param html An HTML string that describes the visual presentation of the content. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         */
        insertHtmlAsSibling(insertLocation: string, html: string): void;
        /**
         *
         * Inserts the image at the specified insert location..
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation The location of the table relative to the current Paragraph.
         * @param base64EncodedImage HTML string to append.
         * @param width Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Inserts the paragraph text at the specifiec insert location.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation The location of the table relative to the current Paragraph.
         * @param paragraphText HTML string to append.
         */
        insertRichTextAsSibling(insertLocation: string, paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns before or after the current paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation The location of the table relative to the current Paragraph.
         * @param rowCount The number of rows in the table.
         * @param columnCount The number of columns in the table.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Paragraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Paragraph;
        toJSON(): {
            "id": string;
            "type": string;
        };
    }
    /**
     *
     * Represents a collection of Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.Paragraph>;
        /**
         *
         * Returns the number of paragraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a Paragraph object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index The ID of the Paragraph object, or the index location of the Paragraph object in the collection.
         */
        getItem(index: number | string): OneNote.Paragraph;
        /**
         *
         * Gets a paragraph on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.ParagraphCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a RichText object in a Paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class RichText extends OfficeExtension.ClientObject {
        private m_id;
        private m_languageId;
        private m_paragraph;
        private m_text;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the Paragraph object that contains the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The language id of the text. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * Gets the text content of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly text: string;
        readonly _ReferenceId: string;
        /**
         *
         * Get the HTML of the rich text
         *
         * [Api set: OneNoteApi]
         * @returns The html of the rich text
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.RichText;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.RichText;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.RichText;
        toJSON(): {
            "id": string;
            "languageId": string;
            "text": string;
        };
    }
    /**
     *
     * Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Image extends OfficeExtension.ClientObject {
        private m_description;
        private m_height;
        private m_hyperlink;
        private m_id;
        private m_ocrData;
        private m_pageContent;
        private m_paragraph;
        private m_width;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets or sets the description of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        description: string;
        /**
         *
         * Gets or sets the height of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        height: number;
        /**
         *
         * Gets or sets the hyperlink of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        hyperlink: string;
        /**
         *
         * Gets the ID of the Image object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly ocrData: OneNote.ImageOcrData;
        /**
         *
         * Gets or sets the width of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        width: number;
        readonly _ReferenceId: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ImageUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Image): void;
        /**
         *
         * Gets the base64-encoded binary representation of the Image.
            Example: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIA...
         *
         * [Api set: OneNoteApi 1.1]
         */
        getBase64Image(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Image;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Image;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Image;
        toJSON(): {
            "description": string;
            "height": number;
            "hyperlink": string;
            "id": string;
            "ocrData": ImageOcrData;
            "width": number;
        };
    }
    /**
     *
     * Represents a table in a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class Table extends OfficeExtension.ClientObject {
        private m_borderVisible;
        private m_columnCount;
        private m_id;
        private m_paragraph;
        private m_rowCount;
        private m_rows;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the Paragraph object that contains the Table object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rows: OneNote.TableRowCollection;
        /**
         *
         * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
         *
         * [Api set: OneNoteApi 1.1]
         */
        borderVisible: boolean;
        /**
         *
         * Gets the number of columns in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly columnCount: number;
        /**
         *
         * Gets the ID of the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the number of rows in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowCount: number;
        readonly _ReferenceId: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Table): void;
        /**
         *
         * Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param values Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         */
        appendColumn(values?: Array<string>): void;
        /**
         *
         * Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param values Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         */
        appendRow(values?: Array<string>): OneNote.TableRow;
        /**
         *
         * Clears the contents of the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets the table cell at a specified row and column.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowIndex The index of the row.
         * @param cellIndex The index of the cell in the row.
         */
        getCell(rowIndex: number, cellIndex: number): OneNote.TableCell;
        /**
         *
         * Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index where the column will be inserted in the table.
         * @param values Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         */
        insertColumn(index: number, values?: Array<string>): void;
        /**
         *
         * Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index where the row will be inserted in the table.
         * @param values Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         */
        insertRow(index: number, values?: Array<string>): OneNote.TableRow;
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Table;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Table;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Table;
        toJSON(): {
            "borderVisible": boolean;
            "columnCount": number;
            "id": string;
            "rowCount": number;
        };
    }
    /**
     *
     * Represents a row in a table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class TableRow extends OfficeExtension.ClientObject {
        private m_cellCount;
        private m_cells;
        private m_id;
        private m_parentTable;
        private m_rowIndex;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cells: OneNote.TableCellCollection;
        /**
         *
         * Gets the parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTable: OneNote.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cellCount: number;
        /**
         *
         * Gets the ID of the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the index of the row in its parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowIndex: number;
        readonly _ReferenceId: string;
        /**
         *
         * Clears the contents of the row.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Inserts a row before or after the current row.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation Where the new rows should be inserted relative to the current row.
         * @param values Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.
         */
        insertRowAsSibling(insertLocation: string, values?: Array<string>): OneNote.TableRow;
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableRow;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableRow;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableRow;
        toJSON(): {
            "cellCount": number;
            "id": string;
            "rowIndex": number;
        };
    }
    /**
     *
     * Contains a collection of TableRow objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class TableRowCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.TableRow>;
        /**
         *
         * Returns the number of table rows in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a table row object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index A number that identifies the index location of a table row object.
         */
        getItem(index: number | string): OneNote.TableRow;
        /**
         *
         * Gets a table row at its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableRowCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableRowCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableRowCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents a cell in a OneNote table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class TableCell extends OfficeExtension.ClientObject {
        private m_cellIndex;
        private m_id;
        private m_paragraphs;
        private m_parentRow;
        private m_rowIndex;
        private m_shadingColor;
        private m__ReferenceId;
        readonly _className: string;
        /**
         *
         * Gets the collection of Paragraph objects in the TableCell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentRow: OneNote.TableRow;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cellIndex: number;
        /**
         *
         * Gets the ID of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowIndex: number;
        /**
         *
         * Gets and sets the shading color of the cell
         *
         * [Api set: OneNoteApi 1.1]
         */
        shadingColor: string;
        readonly _ReferenceId: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableCellUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: TableCell): void;
        /**
         *
         * Adds the specified HTML to the bottom of the TableCell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param html The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param base64EncodedImage HTML string to append.
         * @param width Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param paragraphText HTML string to append.
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowCount Required. The number of rows in the table.
         * @param columnCount Required. The number of columns in the table.
         * @param values Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        appendTable(rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         *
         * Clears the contents of the cell.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableCell;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableCell;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableCell;
        toJSON(): {
            "cellIndex": number;
            "id": string;
            "rowIndex": number;
            "shadingColor": string;
        };
    }
    /**
     *
     * Contains a collection of TableCell objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    class TableCellCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        readonly _className: string;
        /** Gets the loaded child items in this collection. */
        readonly items: Array<OneNote.TableCell>;
        /**
         *
         * Returns the number of tablecells in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        readonly _ReferenceId: string;
        /**
         *
         * Gets a table cell object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index A number that identifies the index location of a table cell object.
         */
        getItem(index: number | string): OneNote.TableCell;
        /**
         *
         * Gets a tablecell at its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableCellCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableCellCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableCellCollection;
        toJSON(): {
            "count": number;
        };
    }
    /**
     *
     * Represents data obtained by OCR (optical character recognition) of an image
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface ImageOcrData {
        /**
         *
         * Represents the OCR language, with values such as EN-US
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrLanguageId: string;
        /**
         *
         * Represents the text obtained by OCR of the image
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrText: string;
    }
    /**
     *
     * Weak reference to an ink stroke object and its content parent
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface InkStrokePointer {
        /**
         *
         * Represents the id of the page content object corresponding to this stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        contentId: string;
        /**
         *
         * Represents the id of the ink stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkStrokeId: string;
    }
    /**
     *
     * Service token for Application::_GetServiceToken.
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface ServiceToken {
        /**
         *
         * Account type
         *
         * [Api set: OneNoteApi 1.1]
         */
        accountType: string;
        /**
         *
         * //
            Header name of the service token
         *
         * [Api set: OneNoteApi 1.1]
         */
        headerName: string;
        /**
         *
         * Header value of the service token
         *
         * [Api set: OneNoteApi 1.1]
         */
        headerValue: string;
    }
    /**
     *
     * Account information.
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface AccountInfo {
        /**
         *
         * Account type
         *
         * [Api set: OneNoteApi 1.1]
         */
        accountType: string;
        /**
         *
         * Account email
         *
         * [Api set: OneNoteApi 1.1]
         */
        email: string;
        /**
         *
         * //
            Account user name
         *
         * [Api set: OneNoteApi 1.1]
         */
        userName: string;
    }
    /**
     *
     * List information for paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface ParagraphInfo {
        /**
         *
         * //
            Bullet list type of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        bulletType: string;
        /**
         *
         * //
            Index of paragraph in list
         *
         * [Api set: OneNoteApi 1.1]
         */
        index: number;
        /**
         *
         * //
            Type of list in paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        listType: string;
        /**
         *
         * //
            number list type of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        numberType: string;
    }
    /**
     *
     * Account information.
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface LoggingInfo {
        /**
         *
         * //
            Correlation Id
         *
         * [Api set: OneNoteApi 1.1]
         */
        correlationId: string;
        /**
         *
         * //
            UI Language
         *
         * [Api set: OneNoteApi 1.1]
         */
        market: string;
        /**
         *
         * //
            Session Id
         *
         * [Api set: OneNoteApi 1.1]
         */
        sessionId: string;
        /**
         *
         * //
            UI Language
         *
         * [Api set: OneNoteApi 1.1]
         */
        uiLanguage: string;
        /**
         *
         * //
            User Id
         *
         * [Api set: OneNoteApi 1.1]
         */
        userId: string;
    }
    /**
     *
     * Account information.
     *
     * [Api set: OneNoteApi 1.1]
     */
    interface LogData {
        /**
         *
         * //
            None PII
         *
         * [Api set: OneNoteApi 1.1]
         */
        isNonPII: boolean;
        /**
         *
         * //
            data tag
         *
         * [Api set: OneNoteApi 1.1]
         */
        tag: string;
        /**
         *
         * //
            data value
         *
         * [Api set: OneNoteApi 1.1]
         */
        value: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace InsertLocation {
        var before: string;
        var after: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace Platform {
        var other: string;
        var web: string;
        var uwp: string;
        var win32: string;
        var mac: string;
        var ios: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace Alignment {
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace Selected {
        var notSelected: string;
        var partialSelected: string;
        var selected: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace PageContentType {
        var outline: string;
        var image: string;
        var ink: string;
        var other: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace ParagraphType {
        var richText: string;
        var image: string;
        var table: string;
        var ink: string;
        var other: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace ServiceId {
        var form: string;
        var entity: string;
        var graph: string;
        var oneService: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace IdentityFilter {
        var selection: string;
        var activeProfile: string;
        var liveId: string;
        var orgId: string;
        var adal: string;
        var notebook: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace ListType {
        var none: string;
        var number: string;
        var bullet: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace AccountType {
        var other: string;
        var liveId: string;
        var orgId: string;
        var adal: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace LogLevel {
        var trace: string;
        var data: string;
        var exception: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace EventFlag {
        /**
         *
         * DefaultEventFlags
         *
         */
        var defaultFlag: string;
        /**
         *
         * CriticalDataEventFlags
         *
         */
        var criticalFlag: string;
        /**
         *
         * MeasureDataEventFlags
         *
         */
        var measureFlag: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace NumberType {
        var none: string;
        var arabic: string;
        var ucroman: string;
        var lcroman: string;
        var ucletter: string;
        var lcletter: string;
        var ordinal: string;
        var cardtext: string;
        var ordtext: string;
        var hex: string;
        var chiManSty: string;
        var dbNum1: string;
        var dbNum2: string;
        var aiueo: string;
        var iroha: string;
        var dbChar: string;
        var sbChar: string;
        var dbNum3: string;
        var dbNum4: string;
        var circlenum: string;
        var darabic: string;
        var daiueo: string;
        var diroha: string;
        var arabicLZ: string;
        var bullet: string;
        var ganada: string;
        var chosung: string;
        var gb1: string;
        var gb2: string;
        var gb3: string;
        var gb4: string;
        var zodiac1: string;
        var zodiac2: string;
        var zodiac3: string;
        var tpeDbNum1: string;
        var tpeDbNum2: string;
        var tpeDbNum3: string;
        var tpeDbNum4: string;
        var chnDbNum1: string;
        var chnDbNum2: string;
        var chnDbNum3: string;
        var chnDbNum4: string;
        var korDbNum1: string;
        var korDbNum2: string;
        var korDbNum3: string;
        var korDbNum4: string;
        var hebrew1: string;
        var arabic1: string;
        var hebrew2: string;
        var arabic2: string;
        var hindi1: string;
        var hindi2: string;
        var hindi3: string;
        var thai1: string;
        var thai2: string;
        var numInDash: string;
        var lcrus: string;
        var ucrus: string;
        var lcgreek: string;
        var ucgreek: string;
        var lim: string;
        var custom: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    namespace ControlId {
        var preinstallClassNotebook: string;
        var distributePageId: string;
        var distributeSection: string;
        var reviewStudentWork: string;
        var openTabForCreateClassNotebook: string;
        var openTabForManageStudent: string;
        var openTabForManageTeacher: string;
        var openTabForGetNotebookLink: string;
        var openTabForTeacherTraining: string;
        var openTabForAddinGuide: string;
        var openTabForEducationBlog: string;
        var openTabForEducatorCommunity: string;
        var openTabToSendFeedback: string;
        var openTabForViewKnowledgeBase: string;
        var openTabForSuggestingFeature: string;
    }
    namespace ErrorCodes {
        var generalException: string;
    }
    module Interfaces {
        /** An interface for updating data on the Page object, for use in "page.set({ ... })". */
        interface PageUpdateData {
            /**
             *
             * Gets or sets the indentation level of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            pageLevel?: number;
            /**
             *
             * Gets or sets the title of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            title?: string;
        }
        /** An interface for updating data on the PageContent object, for use in "pageContent.set({ ... })". */
        interface PageContentUpdateData {
            /**
             *
             * Gets or sets the left (X-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            left?: number;
            /**
             *
             * Gets or sets the top (Y-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            top?: number;
        }
        /** An interface for updating data on the Image object, for use in "image.set({ ... })". */
        interface ImageUpdateData {
            /**
             *
             * Gets or sets the description of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            description?: string;
            /**
             *
             * Gets or sets the height of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            height?: number;
            /**
             *
             * Gets or sets the hyperlink of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            hyperlink?: string;
            /**
             *
             * Gets or sets the width of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the Table object, for use in "table.set({ ... })". */
        interface TableUpdateData {
            /**
             *
             * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
             *
             * [Api set: OneNoteApi 1.1]
             */
            borderVisible?: boolean;
        }
        /** An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })". */
        interface TableCellUpdateData {
            /**
             *
             * Gets and sets the shading color of the cell
             *
             * [Api set: OneNoteApi 1.1]
             */
            shadingColor?: string;
        }
    }
}
declare module OneNote {
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_onenote;
        constructor(url?: string);
        readonly application: Application;
    }
    /**
     * Executes a batch script that performs actions on the OneNote object model, using a new request context. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the OneNote application. Since the Office add-in and the OneNote application run in two different processes, the request context is required to get access to the OneNote object model from the add-in.
     */
    function run<T>(batch: (context: OneNote.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the OneNote object model, using the request context of a previously-created API object.
     * @param object - A previously-created API object. The batch will use the same request context as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    function run<T>(object: OfficeExtension.ClientObject, batch: (context: OneNote.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the OneNote object model, using the request context of previously-created API objects.
     * @param object - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared request context, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: OneNote.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}
