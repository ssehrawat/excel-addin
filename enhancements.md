# Task: Implement below enhancements, commit and push to git repo each enhancement with meanigful and detailed commit message. Mark the completed enhancement with [x] and record its GITSHA next to it and branch too. For e.g.:
    - [x] Build landing page
        - GITSHA: 'xyz1234' main

## Enhancement 1: 
    - [x] Stream data from LLM to the taskpane instead of showing all the messages at the completion of the call like its doing currently. Create a separate branch for this change.
        - GITSHA: 'a2d9202c58e7767ffa78ac27fe203fba7a4f4a6d' feat/streaming-taskpane

## Enhancement 2:
    - [] Sometimes the answer suggests user to use certain formulas in excel but LLM does not put them in cell updates. Instruct LLM to include these formulas in cell updates. Insert actual formulas and only use placeholders as last resort. Don't create a separate branch for it.

## Enhancement 3:
    - [x] Allow Workbook Copilot to create Excel charts based on prompt instructions, including backend schema/provider updates and taskpane chart rendering.
        - GITSHA: '37688b46e57e7e46ee77d67e0afb9f10c6ea83eb' feat/streaming-taskpane