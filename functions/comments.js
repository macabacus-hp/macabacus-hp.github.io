/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */



export async function deleteCommentsAndNotes() {
    try {
        await Excel.run(async context => {
        const comments = context.workbook.comments.load();
        await context.sync();
        comments.items.forEach(comment => {
            comment.delete();
        })
        await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}
  
export async function deleteResolvedComments() {
    try {
        await Excel.run(async context => {
        const comments = context.workbook.comments.load();
        await context.sync();
        comments.items.forEach(comment => {
            comment.resolved === true ? comment.delete() : console.log('not resolved yet')
        })
        await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

export async function resolveComments() {
    try {
        await Excel.run(async context => {
        const comments = context.workbook.comments.load();
        await context.sync();
        comments.items.forEach(comment => {
            comment.resolved === false ? comment.resolved = true : console.log("comment is already resolved")
        })
        await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

export async function reopenComments() {
    try {
        await Excel.run(async context => {
        const comments = context.workbook.comments.load();
        await context.sync();
        comments.items.forEach(comment => {
            comment.resolved === true ? comment.resolved = false : console.log('Comment Already open')
        })
        await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

export async function removeAuthor() {
    try {
        await Excel.run(async context => {
        const comments = context.workbook.comments.load();
        await context.sync();
        comments.items.forEach(comment => {
            console.log(comment.authorName);
            // comment.set({content: ""})
            // comment.resolved === true ? comment.delete() : console.log('not resolved yet')
        })
        await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}