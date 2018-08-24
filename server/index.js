const { ApolloServer, gql } = require ('apollo-server');

const books = [
    {
        title: "Godfather"
    },
    {
        title: "White Tiger"
    }
];

const typeDefs = gql`
    type Book {
        title: String
    }
    type Query {
        books: [Book]
    }
`;

const resolvers = {
    Query: {
        books: (parent,args,context) => {
            console.log(context.authorization);
            return books;
        },
    }
}

const server = new ApolloServer({
    typeDefs,
    resolvers,
    context: ({req})=> ({
        authorization: req.headers.authorization
    })
});

server.listen().then(({url})=>{
    console.log(`Server ready at ${url}`);
});